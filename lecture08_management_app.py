import io, math, hashlib, re
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='Lecture 08 Management Authority', layout='wide')

CONFIG={
 'physical_check_rate':0.20,'max_discrepancy_percent':0.10,'min_discrepancy_units':1,
 'max_po_skus':5,'target_stock':20,'currency':'INR','payment_terms':'Net 30',
 'tracking_prefix':'INKOR-L08-','transfer_prefix':'INKOR-L08-TRF-','po_prefix':'INKOR-L08-PO-',
 'send_customer_notification':True,
}
SUPPLIERS=[
 {'Supplier_ID':'SUP-001','Supplier_Name':'INKOR FMCG Distribution North','Contact_Name':'Amit Verma','Email':'supplier01@example.com','Phone':'+919100000001','Address':'101 Distribution Park','City':'Raipur','State_Province':'Chhattisgarh','Postal_Code':'492001','Country':'India'},
 {'Supplier_ID':'SUP-002','Supplier_Name':'INKOR FMCG Distribution Central','Contact_Name':'Neha Sharma','Email':'supplier02@example.com','Phone':'+919100000002','Address':'202 Wholesale Market','City':'Raipur','State_Province':'Chhattisgarh','Postal_Code':'492001','Country':'India'},
 {'Supplier_ID':'SUP-003','Supplier_Name':'INKOR FMCG Distribution South','Contact_Name':'Rahul Patel','Email':'supplier03@example.com','Phone':'+919100000003','Address':'303 Logistics Park','City':'Raipur','State_Province':'Chhattisgarh','Postal_Code':'492012','Country':'India'},
]
INV_REQ=['Handle','Title','SKU','Location','Incoming (not editable)','Unavailable (not editable)','Committed (not editable)','Available (not editable)','On hand (current)']
PROD_REQ=['Handle','Title','Vendor','Variant SKU','Cost per item']

def read_csv(upload):
    if upload is None:return None
    upload.seek(0)
    try:return pd.read_csv(upload,dtype=str,keep_default_na=False)
    except UnicodeDecodeError:
        upload.seek(0);return pd.read_csv(upload,dtype=str,encoding='latin1',keep_default_na=False)

def require(df, cols, label):
    miss=[c for c in cols if c not in df.columns]
    if miss: raise ValueError(f'{label}: missing required columns: {", ".join(miss)}')

def to_int(v, field):
    s=str(v).strip()
    if s=='': return 0
    try:
        x=float(s)
        if not x.is_integer(): raise ValueError
        return int(x)
    except: raise ValueError(f'Invalid integer in {field}: {v!r}')

def num_inventory(df):
    out=df.copy()
    for c in ['Incoming (not editable)','Unavailable (not editable)','Committed (not editable)','Available (not editable)','On hand (current)']:
        out[c]=[to_int(v,c) for v in out[c]]
    return out

def inventory_check(inv):
    rows=[]
    for _,r in inv.iterrows():
        calc=r['Available (not editable)']+r['Committed (not editable)']+r['Unavailable (not editable)']
        rows.append({'Handle':r['Handle'],'Title':r['Title'],'SKU':r['SKU'],'Location':r['Location'],'Bin_Name':r.get('Bin name',''),
          'Incoming':r['Incoming (not editable)'],'Unavailable':r['Unavailable (not editable)'],'Committed':r['Committed (not editable)'],
          'Available':r['Available (not editable)'],'On_Hand':r['On hand (current)'],'Calculated_On_Hand':calc,
          'Interpretation_Status':'MATCH' if calc==r['On hand (current)'] else 'ERROR'})
    return pd.DataFrame(rows)

def col(df,*names):
    low={str(c).strip().lower():c for c in df.columns}
    for n in names:
        if n.lower() in low:return low[n.lower()]
    return None

def normalize_orders(df):
    # Shopify order CSV field aliases; fields absent from export are not fabricated.
    mapping={
      'Order_Name':col(df,'Name','Order','Order Name'), 'Financial_Status':col(df,'Financial Status','Payment Status'),
      'Fulfillment_Status':col(df,'Fulfillment Status','Fulfilment Status'), 'Source':col(df,'Source','Source Name','Channel'),
      'SKU':col(df,'Lineitem sku','Lineitem SKU','SKU'), 'Product':col(df,'Lineitem name','Lineitem Name','Product','Title'),
      'Order_Quantity':col(df,'Lineitem quantity','Lineitem Quantity','Quantity'),
    }
    essential=['Order_Name','SKU','Order_Quantity']
    miss=[k for k in essential if not mapping[k]]
    if miss: raise ValueError('Orders CSV: cannot identify required Shopify fields: '+', '.join(miss))
    o=pd.DataFrame()
    for k,c in mapping.items(): o[k]=df[c].astype(str) if c else ''
    o['Order_Quantity']=[to_int(v,'Order quantity') for v in o['Order_Quantity']]
    return o

def immediate_fulfilment(orders, inv, products):
    # Conservative: only a SKU whose total available at one location can cover the line is authorized.
    prod=dict(zip(products['Variant SKU'],products['Title']))
    avail={(r['SKU'],r['Location']):r['Available (not editable)'] for _,r in inv.iterrows() if r['SKU']}
    rows=[]; seq=1
    for _,r in orders.sort_values(['Order_Name','SKU']).iterrows():
        fs=str(r['Fulfillment_Status']).lower()
        if fs in ('fulfilled','fulfilled successfully'): continue
        qty=r['Order_Quantity']; sku=r['SKU']
        candidates=sorted([(a,loc) for (s,loc),a in avail.items() if s==sku and a>=qty], key=lambda x:(-x[0],x[1]))
        if not candidates: continue
        a,loc=candidates[0]; avail[(sku,loc)]-=qty
        rows.append({'Action_ID':f'FUL-{seq:03d}','Order_Name':r['Order_Name'],'Financial_Status':r['Financial_Status'],
          'Fulfillment_Status':r['Fulfillment_Status'],'Source':r['Source'],'SKU':sku,'Product':prod.get(sku,r['Product']),
          'Order_Quantity':qty,'Outstanding_Quantity':qty,'Fulfillment_Location':loc,'Available_Before':a,'Quantity_to_Fulfil':qty,
          'Shipping_Carrier':'Course Simulation','Tracking_Number':f"{CONFIG['tracking_prefix']}{seq:04d}",
          'Send_Customer_Notification':CONFIG['send_customer_notification'],
          'Instruction':f'Fulfil {qty} unit(s) of {sku} from {loc}; enter the supplied course tracking details.'})
        seq+=1
    return pd.DataFrame(rows)

def corrections(inv):
    eligible=inv[(inv['SKU'].str.strip()!='') & (inv['On hand (current)']>0)].sort_values(['Location','SKU']).copy()
    n=max(1,math.ceil(len(eligible)*CONFIG['physical_check_rate'])) if len(eligible) else 0
    eligible=eligible.head(n); kinds=['PHYSICAL_SHORTAGE','PHYSICAL_SURPLUS','DAMAGED','QUALITY_CONTROL']; rows=[]
    for i,(_,r) in enumerate(eligible.iterrows(),1):
        kind=kinds[(i-1)%4]; on=r['On hand (current)']; av=r['Available (not editable)']
        d=max(CONFIG['min_discrepancy_units'],math.floor(on*CONFIG['max_discrepancy_percent']))
        if kind in ('PHYSICAL_SHORTAGE','DAMAGED','QUALITY_CONTROL'): d=min(d,max(0,av))
        if d<=0: continue
        physical=on-d if kind=='PHYSICAL_SHORTAGE' else on+d if kind=='PHYSICAL_SURPLUS' else on
        damaged=d if kind in ('DAMAGED','QUALITY_CONTROL') else 0
        action={'PHYSICAL_SHORTAGE':'DECREASE_ON_HAND','PHYSICAL_SURPLUS':'INCREASE_ON_HAND','DAMAGED':'MOVE_TO_UNAVAILABLE','QUALITY_CONTROL':'MOVE_TO_UNAVAILABLE'}[kind]
        expected=physical if kind.startswith('PHYSICAL_') else on
        rows.append({'Action_ID':f'COR-{i:03d}','Handle':r['Handle'],'Title':r['Title'],'SKU':r['SKU'],'Location':r['Location'],
          'Shopify_On_Hand':on,'Shopify_Committed':r['Committed (not editable)'],'Shopify_Unavailable':r['Unavailable (not editable)'],
          'Shopify_Available':av,'Finding_Type':kind,'Reported_Physical_Quantity':physical,'Reported_Damaged_Quantity':damaged,
          'Authorized_Action':action,'Authorized_Quantity':d,'Adjustment_Reason':kind.replace('_',' ').title(),'Expected_On_Hand':expected,
          'Instruction':f'Apply {action} for {d} unit(s) at {r["Location"]} using the stated reason.'})
    return pd.DataFrame(rows)

def transfers(inv, products):
    prod=dict(zip(products['Variant SKU'],products['Title'])); rows=[]; seq=1
    for sku,g in inv[inv['SKU'].str.strip()!=''].groupby('SKU'):
        if len(g)<2: continue
        g=g.sort_values(['Available (not editable)','Location'])
        dest=g.iloc[0]; origin=g.iloc[-1]
        diff=origin['Available (not editable)']-dest['Available (not editable)']
        reserved=origin['Committed (not editable)']; surplus=max(0,origin['Available (not editable)']-reserved)
        qty=min(surplus,max(0,diff//2))
        if qty<1: continue
        rows.append({'Action_ID':f'TRF-{seq:03d}','SKU':sku,'Product':prod.get(sku,''),'Origin_Location':origin['Location'],
          'Destination_Location':dest['Location'],'Origin_Available':origin['Available (not editable)'],'Destination_Available':dest['Available (not editable)'],
          'Transferable_Surplus':surplus,'Quantity_to_Transfer':qty,'Reference_Name':f"{CONFIG['transfer_prefix']}{seq:03d}",
          'Shipping_Carrier':'Course Simulation','Tracking_Number':f'INKOR-TRF-{seq:04d}','Expected_Arrival':date.today()+timedelta(days=2),
          'Receive_Action':'FULL','Quantity_to_Receive':qty,'Instruction':f'Transfer {qty} unit(s) of {sku} from {origin["Location"]} to {dest["Location"]} and receive in full.'})
        seq+=1
        if seq>5: break
    return pd.DataFrame(rows)

def vendor_supplier_map(products):
    vendors=sorted(v for v in products['Vendor'].astype(str).str.strip().unique() if v)
    return {v:SUPPLIERS[i%len(SUPPLIERS)]['Supplier_ID'] for i,v in enumerate(vendors)}

def purchase_order(inv, products):
    p=products.copy(); p['CostNum']=pd.to_numeric(p['Cost per item'],errors='coerce')
    totals=inv.groupby('SKU',as_index=False).agg(Current_Available=('Available (not editable)','sum'),On_Hand=('On hand (current)','sum'))
    m=totals.merge(p[['Variant SKU','Title','Vendor','CostNum']],left_on='SKU',right_on='Variant SKU',how='left')
    m=m[(m['SKU'].str.strip()!='') & m['CostNum'].notna() & (m['CostNum']>0)].copy()
    m['Outstanding_Demand']=0
    m=m.sort_values(['Current_Available','On_Hand','SKU']).head(CONFIG['max_po_skus'])
    vmap=vendor_supplier_map(products); locations=sorted(x for x in inv['Location'].unique() if str(x).strip())
    if not locations: raise ValueError('No inventory locations found.')
    destination=locations[0]; rows=[]; used=set()
    for i,(_,r) in enumerate(m.iterrows(),1):
        sid=vmap.get(str(r['Vendor']).strip())
        if not sid: continue
        used.add(sid); required=max(CONFIG['target_stock'],int(r['Outstanding_Demand']))
        qty=max(1,required-int(r['Current_Available']))
        rows.append({'PO_Action_ID':f'PO-{i:03d}','Supplier_ID':sid,'Supplier_Name':next(s['Supplier_Name'] for s in SUPPLIERS if s['Supplier_ID']==sid),
          'Destination_Location':destination,'Reference_Number':f"{CONFIG['po_prefix']}{i:03d}",'Supplier_Currency':CONFIG['currency'],
          'Payment_Terms':CONFIG['payment_terms'],'SKU':r['SKU'],'Product':r['Title'],'Supplier_SKU':'','Current_Available':int(r['Current_Available']),
          'Outstanding_Demand':0,'Target_Stock':CONFIG['target_stock'],'Quantity':qty,'Unit_Cost':float(r['CostNum']),'Tax_Percent':0.0,
          'Note_to_Supplier':'Lecture 08 course simulation','Instruction':f'Add {qty} unit(s) of {r["SKU"]} to the purchase order.'})
    supplier_df=pd.DataFrame([{**s,'Instruction':'Create/select this course-simulated supplier in Shopify.'} for s in SUPPLIERS if s['Supplier_ID'] in used])
    return supplier_df,pd.DataFrame(rows)

def xlsx_bytes(sheets):
    wb=Workbook(); wb.remove(wb.active)
    for name,df in sheets.items():
        ws=wb.create_sheet(name[:31]); ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:{get_column_letter(max(1,len(df.columns)))}{max(1,len(df)+1)}'
        for j,c in enumerate(df.columns,1):
            cell=ws.cell(1,j,c); cell.font=Font(bold=True); cell.fill=PatternFill('solid',fgColor='D9EAF7'); cell.alignment=Alignment(wrap_text=True)
        for i,row in enumerate(df.itertuples(index=False,name=None),2):
            for j,v in enumerate(row,1): ws.cell(i,j,v)
        for j,c in enumerate(df.columns,1):
            vals=[len(str(c))]+[len(str(x)) for x in df[c].head(100).tolist()]
            ws.column_dimensions[get_column_letter(j)].width=min(max(vals)+2,40)
    b=io.BytesIO(); wb.save(b); return b.getvalue()

def inventory_import_from_po(inv, po):
    # Uses Shopify inventory-export structure; only On hand (new) is changed for PO receipt.
    out=inv.copy()
    if 'On hand (new)' not in out.columns: out['On hand (new)']=''
    for _,r in po.iterrows():
        mask=(out['SKU']==r['SKU']) & (out['Location']==r['Destination_Location'])
        if not mask.any(): continue
        idx=out.index[mask][0]
        out.at[idx,'On hand (new)']=int(out.at[idx,'On hand (current)'])+int(r['Quantity'])
    # restore Shopify-like text columns; current numeric columns are fine in CSV.
    return out

st.title('Lecture 08 — INKOR FMCG Management Authority')
st.caption('Deterministic course-simulation generator. Shopify source values are never silently invented.')

tabs=st.tabs(['1 · Inventory & Order Instructions','2 · Purchase Order Instructions','3 · Replenishment Inventory CSV','4 · Post-Replenishment Order Instructions'])
with tabs[0]:
    st.subheader('01_Inventory_Management_Instructions.xlsx')
    oc=st.file_uploader('Orders Export CSV',type='csv',key='s1o'); ic=st.file_uploader('Inventory Export CSV',type='csv',key='s1i'); pc=st.file_uploader('Products Export CSV',type='csv',key='s1p')
    if st.button('Validate and generate Workbook 01',key='b1'):
      try:
        orders=read_csv(oc); inv=read_csv(ic); products=read_csv(pc)
        if any(x is None for x in [orders,inv,products]): raise ValueError('Upload all three CSV files.')
        require(inv,INV_REQ,'Inventory CSV'); require(products,PROD_REQ,'Products CSV'); inv=num_inventory(inv); no=normalize_orders(orders)
        chk=inventory_check(inv)
        if (chk['Interpretation_Status']=='ERROR').any(): raise ValueError('Inventory state arithmetic failed. Workbook generation stopped.')
        ful=immediate_fulfilment(no,inv,products); cor=corrections(inv); tr=transfers(inv,products)
        data=xlsx_bytes({'Inventory_Check':chk,'Immediate_Order_Fulfilment':ful,'Inventory_Corrections':cor,'Transfer_Actions':tr})
        st.success(f'Validated {len(inv)} inventory rows. Workbook generated.')
        st.download_button('Download 01_Inventory_Management_Instructions.xlsx',data,'01_Inventory_Management_Instructions.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      except Exception as e: st.error(str(e))
with tabs[1]:
    st.subheader('02_Purchase_Order_Instructions.xlsx')
    ic=st.file_uploader('Latest Inventory Export CSV',type='csv',key='s2i'); pc=st.file_uploader('Products Export CSV',type='csv',key='s2p')
    if st.button('Validate and generate Workbook 02',key='b2'):
      try:
        inv=read_csv(ic); products=read_csv(pc)
        if inv is None or products is None: raise ValueError('Upload both CSV files.')
        require(inv,INV_REQ,'Inventory CSV'); require(products,PROD_REQ,'Products CSV'); inv=num_inventory(inv)
        suppliers,po=purchase_order(inv,products)
        if po.empty: raise ValueError('No eligible products with valid cost were found for automatic PO generation.')
        data=xlsx_bytes({'Supplier_Details':suppliers,'Purchase_Order_Details':po})
        st.success(f'Generated PO instructions for {len(po)} SKU(s).')
        st.download_button('Download 02_Purchase_Order_Instructions.xlsx',data,'02_Purchase_Order_Instructions.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      except Exception as e: st.error(str(e))
with tabs[2]:
    st.subheader('03_Inventory_Update.csv')
    ic=st.file_uploader('Latest Inventory Export CSV',type='csv',key='s3i'); po_file=st.file_uploader('02_Purchase_Order_Instructions.xlsx',type='xlsx',key='s3po')
    if st.button('Generate Shopify-compatible Inventory Update CSV',key='b3'):
      try:
        inv=read_csv(ic)
        if inv is None or po_file is None: raise ValueError('Upload latest Inventory CSV and Workbook 02.')
        require(inv,INV_REQ,'Inventory CSV'); inv=num_inventory(inv)
        po=pd.read_excel(po_file,sheet_name='Purchase_Order_Details',dtype=str).fillna('')
        for c in ['SKU','Destination_Location','Quantity']:
            if c not in po.columns: raise ValueError(f'Purchase_Order_Details missing {c}.')
        po['Quantity']=[to_int(v,'PO Quantity') for v in po['Quantity']]
        out=inventory_import_from_po(inv,po)
        data=out.to_csv(index=False).encode('utf-8-sig')
        st.success('Inventory update CSV generated from PO quantities.')
        st.download_button('Download 03_Inventory_Update.csv',data,'03_Inventory_Update.csv','text/csv')
      except Exception as e: st.error(str(e))
with tabs[3]:
    st.subheader('04_Order_Processing_Instructions.xlsx')
    st.info('This stage uses the latest Orders and Inventory exports after replenishment. It generates fulfil/cancel instructions only from actual uploaded orders.')
    oc=st.file_uploader('Latest Orders Export CSV',type='csv',key='s4o'); ic=st.file_uploader('Latest Inventory Export CSV',type='csv',key='s4i'); pc=st.file_uploader('Products Export CSV',type='csv',key='s4p')
    if st.button('Generate Workbook 04',key='b4'):
      try:
        orders=read_csv(oc); inv=read_csv(ic); products=read_csv(pc)
        if any(x is None for x in [orders,inv,products]): raise ValueError('Upload all three CSV files.')
        require(inv,INV_REQ,'Inventory CSV'); require(products,PROD_REQ,'Products CSV'); inv=num_inventory(inv); no=normalize_orders(orders)
        ful=immediate_fulfilment(no,inv,products)
        # deterministic cancellation candidate: first remaining non-fulfilled order not selected for fulfilment
        selected=set(ful['Order_Name'].tolist()) if not ful.empty else set(); cand=no[~no['Order_Name'].isin(selected)].sort_values(['Order_Name','SKU'])
        cancel=[]
        if not cand.empty:
            r=cand.iloc[0]; cancel=[{'Action_ID':'CAN-001','Order_Name':r['Order_Name'],'SKU':r['SKU'],'Product':r['Product'],'Cancel_Instruction':'Cancel order as management-authorized course scenario','Restock_Instruction':'Restock inventory where Shopify presents the option','Refund_Instruction':'Follow Shopify-displayed refund options for the actual order state','Reason':'Management-authorized cancellation exercise'}]
        data=xlsx_bytes({'Fulfil_Pending_Orders':ful,'Cancel_Order':pd.DataFrame(cancel)})
        st.success('Post-replenishment order instructions generated.')
        st.download_button('Download 04_Order_Processing_Instructions.xlsx',data,'04_Order_Processing_Instructions.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      except Exception as e: st.error(str(e))

st.divider(); st.caption('Course simulation only. Review generated instructions before student use. No real supplier, carrier, shipment, or payment event is represented.')
