#!/usr/bin/env python3
"""Streamlit app for Lecture 08 INKOR FMCG management instructions.

Stage 1:
  Orders CSV + Inventory CSV + Products CSV
  -> 01_Inventory_Management_Instructions.xlsx

Stage 2:
  Latest Inventory CSV + Products CSV
  -> 02_Purchase_Order_Instructions.xlsx

No inventory-update CSV is generated.
All Shopify-imported facts are preserved; management-simulated facts are deterministic.
"""
from __future__ import annotations

import hashlib
import math
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CONFIG = {
    "max_inventory_corrections": 5,
    "max_transfers": 5,
    "max_discrepancy_percent": 0.10,
    "min_discrepancy_units": 1,
    "max_po_skus": 5,
    "manual_po_skus": 2,
    "import_po_skus": 3,
    "target_stock": 20,
    "currency": "INR",
    "payment_terms": "Net 30",
    "tracking_prefix": "INKOR-L08-",
    "transfer_prefix": "INKOR-L08-TRF-",
    "po_prefix": "INKOR-L08-PO-",
    "shipping_carrier": "Course Simulated Carrier",
    "send_customer_notification": True,
    "tax_percent": Decimal("0"),  # course configuration; change here if a different fixed value is approved
}

SUPPLIERS = {
    "SUP-001": {"Supplier_Name": "INKOR FMCG Distribution North", "Contact_Name": "Amit Verma", "Email": "supplier01@example.com", "Phone": "+919100000001", "Address": "101 Distribution Park", "City": "Raipur", "State_Province": "Chhattisgarh", "Postal_Code": "492001", "Country": "India"},
    "SUP-002": {"Supplier_Name": "INKOR FMCG Distribution Central", "Contact_Name": "Neha Sharma", "Email": "supplier02@example.com", "Phone": "+919100000002", "Address": "202 Wholesale Market", "City": "Raipur", "State_Province": "Chhattisgarh", "Postal_Code": "492001", "Country": "India"},
    "SUP-003": {"Supplier_Name": "INKOR FMCG Distribution South", "Contact_Name": "Rahul Patel", "Email": "supplier03@example.com", "Phone": "+919100000003", "Address": "303 Logistics Park", "City": "Raipur", "State_Province": "Chhattisgarh", "Postal_Code": "492012", "Country": "India"},
}

INV_REQUIRED = ["Handle","Title","SKU","Location","Bin name","Incoming (not editable)","Unavailable (not editable)","Committed (not editable)","Available (not editable)","On hand (current)"]
PROD_REQUIRED = ["Title","Vendor","Variant SKU","Cost per item"]
ORDER_REQUIRED = ["Name","Financial Status","Fulfillment Status","Lineitem quantity","Lineitem name","Lineitem sku","Lineitem fulfillment status","Location","Source","Cancelled at","Refunded Amount"]

INV_NUM = ["Incoming (not editable)","Unavailable (not editable)","Committed (not editable)","Available (not editable)","On hand (current)"]
ORDER_LEVEL_FFILL = ["Name","Financial Status","Fulfillment Status","Location","Source","Cancelled at","Refunded Amount"]

SHEET1_COLS = {
    "Inventory_Check": ["Handle","Title","SKU","Location","Bin_Name","Incoming","Unavailable","Committed","Available","On_Hand","Calculated_On_Hand","Interpretation_Status"],
    "Immediate_Order_Fulfilment": ["Action_ID","Order_Name","Financial_Status","Fulfillment_Status","Source","SKU","Product","Order_Quantity","Outstanding_Quantity","Fulfillment_Location","Available_Before","Quantity_to_Fulfil","Shipping_Carrier","Tracking_Number","Send_Customer_Notification","Instruction"],
    "Inventory_Corrections": ["Action_ID","Handle","Title","SKU","Location","Shopify_On_Hand","Shopify_Committed","Shopify_Unavailable","Shopify_Available","Finding_Type","Reported_Physical_Quantity","Reported_Damaged_Quantity","Authorized_Action","Authorized_Quantity","Adjustment_Reason","Expected_On_Hand","Instruction"],
    "Transfer_Actions": ["Action_ID","SKU","Product","Origin_Location","Destination_Location","Origin_Available","Destination_Available","Transferable_Surplus","Quantity_to_Transfer","Reference_Name","Shipping_Carrier","Tracking_Number","Expected_Arrival","Receive_Action","Quantity_to_Receive","Instruction"],
}
SHEET2_COLS = {
    "Supplier_Details": ["Supplier_ID","Supplier_Name","Contact_Name","Email","Phone","Address","City","State_Province","Postal_Code","Country","Instruction"],
    "Purchase_Order_Details": ["PO_Action_ID","Supplier_ID","Supplier_Name","Destination_Location","Reference_Number","Supplier_Currency","Payment_Terms","SKU","Product","Supplier_SKU","Current_Available","Outstanding_Demand","Target_Stock","Quantity","Unit_Cost","Tax_Percent","Note_to_Supplier","Instruction"],
}

class DataError(Exception): pass

def read_csv(source) -> pd.DataFrame:
    # Accept a filesystem path or a Streamlit UploadedFile/BytesIO.
    if hasattr(source, "seek"):
        source.seek(0)
    return pd.read_csv(source, dtype=str, keep_default_na=False, encoding="utf-8-sig")

def require(df, cols, label):
    miss = [c for c in cols if c not in df.columns]
    if miss: raise DataError(f"{label}: missing required column(s): {', '.join(miss)}")

def to_int_series(s, label):
    n = pd.to_numeric(s, errors="coerce")
    if n.isna().any():
        bad = s[n.isna()].head(5).tolist()
        raise DataError(f"{label}: non-numeric quantity value(s): {bad}")
    if ((n % 1) != 0).any(): raise DataError(f"{label}: non-integer quantity found")
    return n.astype(int)

def load_inventory(path):
    d = read_csv(path); require(d, INV_REQUIRED, "Inventory CSV")
    d = d.copy(); d["SKU"] = d["SKU"].str.strip(); d["Location"] = d["Location"].str.strip()
    if (d["SKU"] == "").any(): raise DataError("Inventory CSV: blank SKU found")
    if (d["Location"] == "").any(): raise DataError("Inventory CSV: blank Location found")
    if d.duplicated(["SKU","Location"]).any():
        x=d.loc[d.duplicated(["SKU","Location"],False),["SKU","Location"]].head().to_dict("records")
        raise DataError(f"Inventory CSV: duplicate SKU+Location mapping: {x}")
    for c in INV_NUM: d[c] = to_int_series(d[c], f"Inventory CSV/{c}")
    if (d[["Incoming (not editable)","Unavailable (not editable)","Committed (not editable)"]] < 0).any().any(): raise DataError("Inventory CSV: Incoming/Unavailable/Committed cannot be negative")
    d["Calculated_On_Hand"] = d["Available (not editable)"] + d["Committed (not editable)"] + d["Unavailable (not editable)"]
    d["Interpretation_Status"] = d.apply(lambda r: "MATCH" if r["Calculated_On_Hand"] == r["On hand (current)"] else "ERROR", axis=1)
    return d

def load_products(path):
    d=read_csv(path); require(d, PROD_REQUIRED, "Products CSV")
    d=d.copy()
    # Shopify product exports may show product-level fields only on the first variant row.
    # Propagate Title/Vendor only within the same Handle; never across products.
    if "Handle" in d.columns:
        d["Title"] = d.groupby("Handle", sort=False)["Title"].transform(lambda x: x.replace("", pd.NA).ffill().bfill()).fillna("")
        d["Vendor"] = d.groupby("Handle", sort=False)["Vendor"].transform(lambda x: x.replace("", pd.NA).ffill().bfill()).fillna("")
    d["Variant SKU"]=d["Variant SKU"].str.strip()
    d=d[d["Variant SKU"]!=""].copy()
    if d.duplicated(["Variant SKU"]).any():
        x=d.loc[d.duplicated(["Variant SKU"],False),["Variant SKU","Title"]].head().to_dict("records")
        raise DataError(f"Products CSV: duplicate/ambiguous Variant SKU mapping: {x}")
    return d

def load_orders(path):
    d=read_csv(path); require(d, ORDER_REQUIRED, "Orders CSV")
    d=d.copy()
    # Shopify exports order-level values on the first line-item row only.
    # First recover the order name, then copy each order's FIRST ROW value to its own line items only.
    d["Name"] = d["Name"].replace("", pd.NA).ffill().fillna("")
    if (d["Name"] == "").any(): raise DataError("Orders CSV: line item found before an order Name")
    for c in [x for x in ORDER_LEVEL_FFILL if x != "Name"]:
        first_by_order = d.groupby("Name", sort=False)[c].transform("first")
        d[c] = first_by_order.fillna("")
    d["Lineitem sku"]=d["Lineitem sku"].str.strip()
    d["Lineitem quantity"]=pd.to_numeric(d["Lineitem quantity"],errors="coerce")
    d=d[d["Lineitem sku"]!=""].copy()  # exclude service/non-inventory lines
    if d["Lineitem quantity"].isna().any(): raise DataError("Orders CSV: non-numeric Lineitem quantity on inventory SKU")
    d["Lineitem quantity"]=d["Lineitem quantity"].astype(int)
    return d

def product_maps(prod):
    title=dict(zip(prod["Variant SKU"],prod["Title"])); vendor=dict(zip(prod["Variant SKU"],prod["Vendor"])); cost=dict(zip(prod["Variant SKU"],prod["Cost per item"]))
    return title,vendor,cost

def valid_order_lines(orders):
    d=orders.copy()
    # Cancelled orders and fully fulfilled line items are not outstanding inventory demand.
    d=d[d["Cancelled at"].str.strip()==""]
    d=d[~d["Lineitem fulfillment status"].str.lower().isin(["fulfilled"])]
    # Fully refunded orders are excluded. Partially refunded remains because line-level refunded qty is not available in export.
    d=d[d["Financial Status"].str.lower().ne("refunded")]
    return d

def order_demand(orders):
    d=valid_order_lines(orders)
    g=(d.groupby(["Name","Financial Status","Fulfillment Status","Source","Location","Lineitem sku","Lineitem name"],dropna=False,as_index=False)["Lineitem quantity"].sum())
    g=g.rename(columns={"Lineitem quantity":"Outstanding_Quantity","Lineitem sku":"SKU","Lineitem name":"Product","Name":"Order_Name","Financial Status":"Financial_Status","Fulfillment Status":"Fulfillment_Status"})
    g["Order_Quantity"]=g["Outstanding_Quantity"]
    return g

def order_sort_key(name):
    s=str(name).lstrip("#")
    digits="".join(ch for ch in s if ch.isdigit())
    return (int(digits) if digits else 10**18, str(name))

def build_inventory_check(inv):
    return pd.DataFrame({
        "Handle":inv["Handle"],"Title":inv["Title"],"SKU":inv["SKU"],"Location":inv["Location"],"Bin_Name":inv["Bin name"],
        "Incoming":inv["Incoming (not editable)"],"Unavailable":inv["Unavailable (not editable)"],"Committed":inv["Committed (not editable)"],"Available":inv["Available (not editable)"],"On_Hand":inv["On hand (current)"],"Calculated_On_Hand":inv["Calculated_On_Hand"],"Interpretation_Status":inv["Interpretation_Status"]})

def build_immediate(orders, inv, prod):
    title,_,_=product_maps(prod); demand=order_demand(orders)
    inv_ok=inv[inv["Interpretation_Status"]=="MATCH"].copy()
    avail={(r.SKU,r.Location):int(getattr(r,"_4")) for r in []}  # placeholder to keep lint simple
    avail={(r["SKU"],r["Location"]):int(r["Available (not editable)"]) for _,r in inv_ok.iterrows()}
    rows=[]; counter=1
    demand=demand.assign(_sort=demand["Order_Name"].map(order_sort_key)).sort_values(["_sort","SKU"])
    for _,r in demand.iterrows():
        sku=r["SKU"]; requested_loc=str(r["Location"]).strip(); qty=int(r["Outstanding_Quantity"])
        if sku not in title or qty<=0: continue
        # If Shopify export provides a valid fulfilment location, preserve it. Otherwise management
        # deterministically authorizes the valid inventory location with greatest Available stock.
        if requested_loc and (sku,requested_loc) in avail:
            loc=requested_loc
        else:
            locs=[(a,l) for (s,l),a in avail.items() if s==sku and a>=qty]
            if not locs: continue
            loc=sorted(locs,key=lambda x:(-x[0],x[1]))[0][1]
        before=avail[(sku,loc)]
        if before >= qty:
            track=f"{CONFIG['tracking_prefix']}{counter:04d}"
            rows.append({"Action_ID":f"FUL-{counter:03d}","Order_Name":r["Order_Name"],"Financial_Status":r["Financial_Status"],"Fulfillment_Status":r["Fulfillment_Status"],"Source":r["Source"],"SKU":sku,"Product":title[sku],"Order_Quantity":int(r["Order_Quantity"]),"Outstanding_Quantity":qty,"Fulfillment_Location":loc,"Available_Before":before,"Quantity_to_Fulfil":qty,"Shipping_Carrier":CONFIG["shipping_carrier"],"Tracking_Number":track,"Send_Customer_Notification":CONFIG["send_customer_notification"],"Instruction":f"Fulfil {qty} unit(s) of {sku} for order {r['Order_Name']} from {loc}; use course-simulated tracking {track}."})
            avail[(sku,loc)] -= qty
            counter += 1
    return pd.DataFrame(rows,columns=SHEET1_COLS["Immediate_Order_Fulfilment"])

def deterministic_corrections(inv):
    """Create at most five sparse, deterministic physical findings.

    Rules:
    - one correction per SKU;
    - spread corrections across locations as evenly as source data permits;
    - cycle through shortage, surplus, damage and quality-control scenarios;
    - never manufacture an impossible damage/QC action where Available is zero.
    """
    eligible=inv[(inv["Interpretation_Status"]=="MATCH") & (inv["On hand (current)"]>0)].copy()
    if eligible.empty:
        return pd.DataFrame(columns=SHEET1_COLS["Inventory_Corrections"])

    eligible=eligible.sort_values(["Location","SKU"]).copy()
    locations=sorted(eligible["Location"].unique())
    # Round-robin by location: 1 here, 1 there, then return to locations if needed.
    by_loc={loc:list(eligible[eligible["Location"]==loc].sort_values("SKU").to_dict("records")) for loc in locations}
    cursors={loc:0 for loc in locations}
    chosen=[]; used_skus=set()
    while len(chosen) < CONFIG["max_inventory_corrections"]:
        progressed=False
        for loc in locations:
            records=by_loc[loc]
            while cursors[loc] < len(records):
                r=records[cursors[loc]]; cursors[loc]+=1
                if r["SKU"] in used_skus: continue
                chosen.append(r); used_skus.add(r["SKU"]); progressed=True
                break
            if len(chosen) >= CONFIG["max_inventory_corrections"]: break
        if not progressed: break

    types=["PHYSICAL_SHORTAGE","PHYSICAL_SURPLUS","DAMAGED","QUALITY_CONTROL"]
    rows=[]
    for r in chosen:
        # Choose the next scenario that is mathematically possible for this row.
        preferred=types[len(rows)%len(types)]
        trial=types[types.index(preferred):]+types[:types.index(preferred)]
        on=int(r["On hand (current)"]); av=int(r["Available (not editable)"])
        disc=max(CONFIG["min_discrepancy_units"],math.floor(on*CONFIG["max_discrepancy_percent"]))
        disc=min(disc,on)
        made=False
        for typ in trial:
            reported=on; damaged=0; expected=on
            if typ=="PHYSICAL_SHORTAGE":
                qty=disc; reported=on-qty; expected=reported; action="ADJUST_ON_HAND_DOWN"; reason="Physical count shortage"
            elif typ=="PHYSICAL_SURPLUS":
                qty=disc; reported=on+qty; expected=reported; action="ADJUST_ON_HAND_UP"; reason="Physical count surplus"
            else:
                if av<=0: continue
                qty=min(disc,av); damaged=qty if typ=="DAMAGED" else 0; action="MOVE_TO_UNAVAILABLE"; reason="Damaged" if typ=="DAMAGED" else "Quality control"
            if qty<=0 or expected<0: continue
            rows.append({"Action_ID":f"COR-{len(rows)+1:03d}","Handle":r["Handle"],"Title":r["Title"],"SKU":r["SKU"],"Location":r["Location"],"Shopify_On_Hand":on,"Shopify_Committed":int(r["Committed (not editable)"]),"Shopify_Unavailable":int(r["Unavailable (not editable)"]),"Shopify_Available":av,"Finding_Type":typ,"Reported_Physical_Quantity":reported,"Reported_Damaged_Quantity":damaged,"Authorized_Action":action,"Authorized_Quantity":qty,"Adjustment_Reason":reason,"Expected_On_Hand":expected,"Instruction":correction_instruction(typ,r["SKU"],r["Location"],qty,reported)})
            made=True; break
        if not made: continue
    return pd.DataFrame(rows,columns=SHEET1_COLS["Inventory_Corrections"])

def correction_instruction(typ,sku,loc,qty,reported):
    if typ=="PHYSICAL_SHORTAGE": return f"Physical count reports {reported} on hand for {sku} at {loc}; reduce on-hand by {qty}."
    if typ=="PHYSICAL_SURPLUS": return f"Physical count reports {reported} on hand for {sku} at {loc}; increase on-hand by {qty}."
    if typ=="DAMAGED": return f"Mark {qty} available unit(s) of {sku} at {loc} as unavailable due to damage."
    return f"Move {qty} available unit(s) of {sku} at {loc} to unavailable for quality control."

def effective_after_corrections(inv,corr):
    e=inv.copy()
    for _,c in corr.iterrows():
        mask=(e["SKU"]==c["SKU"])&(e["Location"]==c["Location"])
        if c["Finding_Type"]=="PHYSICAL_SHORTAGE": e.loc[mask,"Available (not editable)"] -= int(c["Authorized_Quantity"]); e.loc[mask,"On hand (current)"] -= int(c["Authorized_Quantity"])
        elif c["Finding_Type"]=="PHYSICAL_SURPLUS": e.loc[mask,"Available (not editable)"] += int(c["Authorized_Quantity"]); e.loc[mask,"On hand (current)"] += int(c["Authorized_Quantity"])
        else: e.loc[mask,"Available (not editable)"] -= int(c["Authorized_Quantity"]); e.loc[mask,"Unavailable (not editable)"] += int(c["Authorized_Quantity"])
    return e

def build_transfers(inv,corr,prod):
    """Create at most five realistic location transfers.

    A SKU used in Inventory_Corrections is excluded completely from transfers.
    Transfer destinations are spread across locations where the source data permits.
    """
    title,_,_=product_maps(prod)
    e=effective_after_corrections(inv,corr)
    excluded_skus=set(corr["SKU"].astype(str)) if not corr.empty else set()
    candidates=[]
    for sku,g in e[e["Interpretation_Status"]=="MATCH"].groupby("SKU"):
        if sku in excluded_skus or sku not in title or len(g)<2: continue
        for _,dest in g.iterrows():
            dest_av=int(dest["Available (not editable)"])
            shortage=max(0,CONFIG["target_stock"]-dest_av)
            if shortage<=0: continue
            for _,org in g.iterrows():
                if org["Location"]==dest["Location"]: continue
                org_av=int(org["Available (not editable)"])
                # Available already excludes Committed; retain TARGET_STOCK at origin.
                surplus=max(0,org_av-CONFIG["target_stock"])
                if surplus<=0: continue
                candidates.append({"shortage":shortage,"surplus":surplus,"sku":sku,"org":org,"dest":dest})

    candidates.sort(key=lambda x:(-x["shortage"],-x["surplus"],x["sku"],x["dest"]["Location"],x["org"]["Location"]))
    rows=[]; used_skus=set(); used_origin_qty={}; destination_counts={}
    while len(rows) < CONFIG["max_transfers"]:
        remaining=[c for c in candidates if c["sku"] not in used_skus]
        if not remaining: break
        # Prefer a destination used fewer times, then preserve the frozen shortage/surplus priority.
        remaining.sort(key=lambda c:(destination_counts.get(c["dest"]["Location"],0),-c["shortage"],-c["surplus"],c["sku"],c["dest"]["Location"],c["org"]["Location"]))
        c=remaining[0]; sku=c["sku"]; org=c["org"]; dest=c["dest"]
        ok=(sku,org["Location"]); already=used_origin_qty.get(ok,0)
        transferable=max(0,c["surplus"]-already); qty=min(c["shortage"],transferable)
        used_skus.add(sku)
        if qty<1: continue
        idx=len(rows)+1
        # One deterministic partial-receipt example only when enough quantity exists.
        receive_action="PARTIAL" if idx==CONFIG["max_transfers"] and qty>1 else "FULL"
        receive_qty=qty-1 if receive_action=="PARTIAL" else qty
        ref=f"{CONFIG['transfer_prefix']}{idx:03d}"; track=f"{CONFIG['tracking_prefix']}TRF{idx:04d}"
        rows.append({"Action_ID":f"TRF-{idx:03d}","SKU":sku,"Product":title[sku],"Origin_Location":org["Location"],"Destination_Location":dest["Location"],"Origin_Available":int(org["Available (not editable)"]),"Destination_Available":int(dest["Available (not editable)"]),"Transferable_Surplus":transferable,"Quantity_to_Transfer":qty,"Reference_Name":ref,"Shipping_Carrier":CONFIG["shipping_carrier"],"Tracking_Number":track,"Expected_Arrival":date.today()+timedelta(days=2),"Receive_Action":receive_action,"Quantity_to_Receive":receive_qty,"Instruction":f"Transfer {qty} unit(s) of {sku} from {org['Location']} to {dest['Location']}; receive {receive_qty} unit(s) in this exercise."})
        used_origin_qty[ok]=already+qty
        destination_counts[dest["Location"]]=destination_counts.get(dest["Location"],0)+1
    return pd.DataFrame(rows,columns=SHEET1_COLS["Transfer_Actions"])

def vendor_supplier_map(prod):
    vendors=sorted(v.strip() for v in prod["Vendor"].unique() if v.strip())
    ids=list(SUPPLIERS)
    return {v:ids[i%len(ids)] for i,v in enumerate(vendors)}

def decimal_or_none(x):
    try:
        v=Decimal(str(x).strip()); return v if v>0 else None
    except (InvalidOperation,ValueError): return None

def build_po(inv,prod):
    """Build exactly two course POs from the five lowest-inventory eligible SKUs.

    PO-001: first 2 selected SKUs, entered manually in Shopify.
    PO-002: next 3 selected SKUs, added using Shopify's PO product CSV import.

    A Shopify PO has one supplier and one destination. Therefore each PO receives one
    deterministic course supplier and one deterministic destination. Products remain
    the five selected SKUs; quantities are recalculated from the current Available
    inventory of the PO's destination location.
    """
    title,vendor,cost=product_maps(prod)
    ok=inv[inv["Interpretation_Status"]=="MATCH"].copy()
    candidates=[]; excluded=[]
    for sku,g in ok.groupby("SKU"):
        if sku not in title: continue
        c=decimal_or_none(cost.get(sku,""))
        if c is None:
            excluded.append(f"{sku}: missing/zero Cost per item")
            continue
        sg=g.sort_values(["Available (not editable)","On hand (current)","Location"])
        low=sg.iloc[0]
        candidates.append((int(low["Available (not editable)"]),int(low["On hand (current)"]),sku,c))
    candidates.sort(key=lambda x:(x[0],x[1],x[2]))
    selected=candidates[:CONFIG["max_po_skus"]]
    if len(selected) < CONFIG["max_po_skus"]:
        raise DataError(f"Need {CONFIG['max_po_skus']} eligible SKUs for the two-PO exercise; found {len(selected)}.")

    # Two separate POs. Supplier is a PO-level course-management decision and is kept
    # separate from Shopify Product Vendor. Suppliers may be the same or different;
    # this deterministic exercise uses two distinct fixed suppliers.
    po_defs = [
        {"po_id":"PO-001", "reference":f"{CONFIG['po_prefix']}001", "method":"MANUAL", "supplier_id":"SUP-001", "items":selected[:CONFIG["manual_po_skus"]]},
        {"po_id":"PO-002", "reference":f"{CONFIG['po_prefix']}002", "method":"CSV_IMPORT", "supplier_id":"SUP-002", "items":selected[CONFIG["manual_po_skus"]:]},
    ]

    po=[]; used_sup=set(); import_rows=[]
    locations=sorted(ok["Location"].dropna().astype(str).str.strip().loc[lambda x:x!=""].unique())
    if not locations: raise DataError("No valid inventory location available for Purchase Orders.")

    for pdx,pdef in enumerate(po_defs):
        sid=pdef["supplier_id"]; sup=SUPPLIERS[sid]; used_sup.add(sid)
        # Choose one valid PO destination deterministically: the location having the
        # lowest combined Available inventory for this PO's selected SKUs.
        item_skus=[x[2] for x in pdef["items"]]
        scores=[]
        for loc in locations:
            total=0; complete=True
            for sku in item_skus:
                m=ok[(ok["SKU"]==sku)&(ok["Location"]==loc)]
                if m.empty: complete=False; break
                total += int(m.iloc[0]["Available (not editable)"])
            if complete: scores.append((total,loc))
        if not scores:
            raise DataError(f"{pdef['po_id']}: selected SKUs do not share a valid Shopify destination location.")
        dest=sorted(scores,key=lambda x:(x[0],x[1]))[0][1]

        for _,_,sku,costv in pdef["items"]:
            m=ok[(ok["SKU"]==sku)&(ok["Location"]==dest)]
            avail=int(m.iloc[0]["Available (not editable)"])
            outstanding=0
            required=CONFIG["target_stock"]
            qty=max(1,required-avail)
            if avail>=required:
                # The SKU was selected globally as low stock, but this PO destination
                # already meets target; keep the exercise valid by stopping instead of inventing qty.
                raise DataError(f"{pdef['po_id']}/{sku}: destination {dest} already has Available {avail} >= Target {required}.")
            method_text="enter this line manually" if pdef["method"]=="MANUAL" else "import this line using the generated Shopify PO CSV"
            po.append({"PO_Action_ID":pdef["po_id"],"Supplier_ID":sid,"Supplier_Name":sup["Supplier_Name"],"Destination_Location":dest,"Reference_Number":pdef["reference"],"Supplier_Currency":CONFIG["currency"],"Payment_Terms":CONFIG["payment_terms"],"SKU":sku,"Product":title[sku],"Supplier_SKU":"","Current_Available":avail,"Outstanding_Demand":outstanding,"Target_Stock":CONFIG["target_stock"],"Quantity":qty,"Unit_Cost":float(costv),"Tax_Percent":"","Note_to_Supplier":f"Course-simulated Lecture 08 replenishment; {pdef['method']} product-entry exercise.","Instruction":f"Create {pdef['po_id']} for {dest}; {method_text}. Verify SKU, quantity and cost; leave Supplier SKU and Tax blank unless separately authorized; then progress Draft to Ordered."})
            if pdef["method"]=="CSV_IMPORT":
                import_rows.append({"SKU":sku,"Barcode":"","Supplier SKU":"","Quantity":qty,"Cost":str(costv),"Tax":""})

    suppliers=[]
    for sid in sorted(used_sup):
        x=SUPPLIERS[sid]
        suppliers.append({"Supplier_ID":sid,**x,"Instruction":f"Use this course-simulated supplier for the assigned Purchase Order; do not substitute Product Vendor for PO Supplier."})
    import_df=pd.DataFrame(import_rows,columns=["SKU","Barcode","Supplier SKU","Quantity","Cost","Tax"])
    return pd.DataFrame(suppliers,columns=SHEET2_COLS["Supplier_Details"]),pd.DataFrame(po,columns=SHEET2_COLS["Purchase_Order_Details"]),import_df,excluded

def write_xlsx(path, sheets):
    wb=Workbook(); wb.remove(wb.active)
    for name,df in sheets.items():
        ws=wb.create_sheet(name)
        for j,col in enumerate(df.columns,1):
            cell=ws.cell(1,j,col); cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="1F4E78"); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        for i,row in enumerate(df.itertuples(index=False,name=None),2):
            for j,val in enumerate(row,1):
                ws.cell(i,j,val)
        ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
        for j,col in enumerate(df.columns,1):
            vals=[str(col)]+[str(x) for x in df.iloc[:,j-1].head(200).tolist()]
            width=min(max(max(map(len,vals))+2,10),42); ws.column_dimensions[get_column_letter(j)].width=width
        for row in ws.iter_rows():
            for c in row: c.alignment=Alignment(vertical="top",wrap_text=True)
    wb.save(path)



from io import BytesIO
import streamlit as st


def xlsx_bytes(sheets):
    buffer = BytesIO()
    wb=Workbook(); wb.remove(wb.active)
    for name,df in sheets.items():
        ws=wb.create_sheet(name)
        for j,col in enumerate(df.columns,1):
            cell=ws.cell(1,j,col); cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="1F4E78"); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        for i,row in enumerate(df.itertuples(index=False,name=None),2):
            for j,val in enumerate(row,1):
                ws.cell(i,j,val)
        ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
        for j,col in enumerate(df.columns,1):
            vals=[str(col)]+[str(x) for x in df.iloc[:,j-1].head(200).tolist()]
            width=min(max(max(map(len,vals))+2,10),42); ws.column_dimensions[get_column_letter(j)].width=width
        for row in ws.iter_rows():
            for c in row: c.alignment=Alignment(vertical="top",wrap_text=True)
    wb.save(buffer)
    return buffer.getvalue()


def csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


def reset_uploaded(*files):
    for f in files:
        if f is not None and hasattr(f, "seek"):
            f.seek(0)


def run_stage1(orders_file, inventory_file, products_file):
    reset_uploaded(orders_file, inventory_file, products_file)
    inv=load_inventory(inventory_file)
    prod=load_products(products_file)
    orders=load_orders(orders_file)
    invcheck=build_inventory_check(inv)
    imm=build_immediate(orders,inv,prod)
    corr=deterministic_corrections(inv)
    trf=build_transfers(inv,corr,prod)
    overlap=set(corr["SKU"].astype(str)) & set(trf["SKU"].astype(str)) if not corr.empty and not trf.empty else set()
    if overlap:
        raise DataError(f"Internal validation failed: correction/transfer SKU overlap: {sorted(overlap)}")
    if len(corr)>CONFIG["max_inventory_corrections"]:
        raise DataError("Internal validation failed: too many inventory corrections generated.")
    if len(trf)>CONFIG["max_transfers"]:
        raise DataError("Internal validation failed: too many transfers generated.")
    sheets={"Inventory_Check":invcheck,"Immediate_Order_Fulfilment":imm,"Inventory_Corrections":corr,"Transfer_Actions":trf}
    return xlsx_bytes(sheets), sheets


def run_stage2(inventory_file, products_file):
    reset_uploaded(inventory_file, products_file)
    inv=load_inventory(inventory_file)
    prod=load_products(products_file)
    suppliers,po,po_import,excluded=build_po(inv,prod)
    if po.empty:
        raise DataError("No eligible PO rows could be generated from the supplied latest Inventory + Products files.")
    if po["PO_Action_ID"].nunique()!=2:
        raise DataError("Internal validation failed: Stage 2 must generate exactly two Purchase Orders.")
    sheets={"Supplier_Details":suppliers,"Purchase_Order_Details":po}
    return xlsx_bytes(sheets), csv_bytes(po_import), sheets, po_import, excluded


def main_streamlit():
    st.set_page_config(page_title="INKOR FMCG — Lecture 08 Management", page_icon="📦", layout="wide")
    st.title("INKOR FMCG — Lecture 08 Management Application")
    st.caption("Generate deterministic management instructions from Shopify exports. No inventory-update CSV is generated.")

    tab1, tab2 = st.tabs(["Stage 1 — Inventory Management", "Stage 2 — Purchase Orders"])

    with tab1:
        st.subheader("Generate 01_Inventory_Management_Instructions.xlsx")
        st.write("Upload the Shopify Orders, Inventory and Products CSV exports from the same operating state.")
        c1,c2,c3=st.columns(3)
        with c1: orders=st.file_uploader("Orders CSV", type=["csv"], key="s1_orders")
        with c2: inventory=st.file_uploader("Inventory CSV", type=["csv"], key="s1_inventory")
        with c3: products=st.file_uploader("Products CSV", type=["csv"], key="s1_products")
        if st.button("Generate Stage 1 Instructions", type="primary", key="run_s1", disabled=not all([orders,inventory,products])):
            try:
                data,sheets=run_stage1(orders,inventory,products)
                st.session_state["s1_xlsx"]=data
                st.session_state["s1_counts"]={k:len(v) for k,v in sheets.items()}
                st.success("Stage 1 instructions generated successfully.")
            except (DataError,ValueError,KeyError,TypeError) as e:
                st.session_state.pop("s1_xlsx",None); st.session_state.pop("s1_counts",None)
                st.error(str(e))
        if "s1_xlsx" in st.session_state:
            counts=st.session_state["s1_counts"]
            a,b,c,d=st.columns(4)
            a.metric("Inventory Check",counts["Inventory_Check"])
            b.metric("Immediate Fulfilment",counts["Immediate_Order_Fulfilment"])
            c.metric("Corrections",counts["Inventory_Corrections"])
            d.metric("Transfers",counts["Transfer_Actions"])
            st.download_button("Download 01_Inventory_Management_Instructions.xlsx",st.session_state["s1_xlsx"],"01_Inventory_Management_Instructions.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_s1")
            st.caption("Rules enforced: maximum 5 corrections, maximum 5 transfers, one correction/transfer per SKU, and corrected SKUs are excluded from transfers.")

    with tab2:
        st.subheader("Generate 02_Purchase_Order_Instructions.xlsx and PO-002 import CSV")
        st.write("After completing fulfilment, corrections and transfers in Shopify, export the latest Inventory and upload it with Products.")
        c1,c2=st.columns(2)
        with c1: inventory2=st.file_uploader("Latest Inventory CSV", type=["csv"], key="s2_inventory")
        with c2: products2=st.file_uploader("Products CSV", type=["csv"], key="s2_products")
        if st.button("Generate Stage 2 Instructions", type="primary", key="run_s2", disabled=not all([inventory2,products2])):
            try:
                xlsx,csvdata,sheets,po_import,excluded=run_stage2(inventory2,products2)
                st.session_state["s2_xlsx"]=xlsx; st.session_state["s2_csv"]=csvdata
                st.session_state["s2_counts"]={k:len(v) for k,v in sheets.items()}
                st.session_state["s2_import_count"]=len(po_import); st.session_state["s2_excluded"]=excluded
                st.success("Stage 2 instructions and Shopify PO import CSV generated successfully.")
            except (DataError,ValueError,KeyError,TypeError) as e:
                for k in ["s2_xlsx","s2_csv","s2_counts","s2_import_count","s2_excluded"]: st.session_state.pop(k,None)
                st.error(str(e))
        if "s2_xlsx" in st.session_state:
            counts=st.session_state["s2_counts"]
            a,b,c=st.columns(3)
            a.metric("Suppliers",counts["Supplier_Details"])
            b.metric("PO Product Lines",counts["Purchase_Order_Details"])
            c.metric("PO-002 CSV Lines",st.session_state["s2_import_count"])
            st.download_button("Download 02_Purchase_Order_Instructions.xlsx",st.session_state["s2_xlsx"],"02_Purchase_Order_Instructions.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_s2_xlsx")
            st.download_button("Download 02_PO2_Product_Import.csv",st.session_state["s2_csv"],"02_PO2_Product_Import.csv","text/csv",key="dl_s2_csv")
            st.caption("PO-001 contains 2 products for manual entry. PO-002 contains 3 different products for Shopify CSV import.")
            excluded=st.session_state.get("s2_excluded",[])
            if excluded:
                with st.expander("Excluded PO candidates"):
                    for item in excluded: st.write("-",item)


if __name__ == "__main__":
    main_streamlit()
