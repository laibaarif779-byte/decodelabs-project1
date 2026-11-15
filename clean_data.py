import sys, json, pandas as pd, re
SRC, OUT = sys.argv[1], sys.argv[2]
raw = pd.read_excel(SRC, dtype=object)
df = raw.copy()
log = {}
log['rows_in'] = len(raw)
log['nulls_before'] = {k:int(v) for k,v in raw.isna().sum().items() if v}
log['dup_rows_before'] = int(raw.duplicated().sum())
log['dup_ids_before'] = int(raw.OrderID.duplicated().sum())

# text: trim + proper case (IDs upper)
for c in ['Product','PaymentMethod','OrderStatus','ReferralSource','ShippingAddress']:
    df[c] = df[c].astype(str).str.strip().str.replace(r'\s+',' ',regex=True).str.title()
for c in ['OrderID','CustomerID','TrackingNumber']:
    df[c] = df[c].astype(str).str.strip().str.upper()
# coupon: blank = no coupon used (not missing-at-random) -> explicit category, not mode-imputed
df['CouponCode'] = df['CouponCode'].where(df['CouponCode'].notna(), 'No Coupon')
df['CouponCode'] = df['CouponCode'].astype(str).str.strip().str.upper().replace({'NO COUPON':'No Coupon'})
# types
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
for c in ['Quantity','ItemsInCart']: df[c] = pd.to_numeric(df[c]).astype('int64')
for c in ['UnitPrice','TotalPrice']: df[c] = pd.to_numeric(df[c]).round(2)
# duplicates
df = df.drop_duplicates().drop_duplicates('OrderID', keep='first').reset_index(drop=True)

# verification
v = {}
v['rows_out'] = len(df)
v['nulls_after'] = int(df.isna().sum().sum())
v['dup_order_ids'] = int(df.OrderID.duplicated().sum())
v['dup_tracking'] = int(df.TrackingNumber.duplicated().sum())
v['invalid_order_id'] = int((~df.OrderID.str.match(r'^ORD\d{6}$')).sum())
v['invalid_dates'] = int(df.Date.isna().sum())
v['future_dates'] = int((df.Date > pd.Timestamp('2026-11-15')).sum())
v['date_str_bad_format'] = int((~df.Date.dt.strftime('%Y-%m-%d').str.match(r'^\d{4}-\d{2}-\d{2}$')).sum())
v['total_mismatch'] = int(((df.Quantity*df.UnitPrice).round(2) != df.TotalPrice).sum())
v['nonpositive_qty_price'] = int(((df.Quantity<=0)|(df.UnitPrice<=0)).sum())
v['date_min'] = str(df.Date.min().date()); v['date_max'] = str(df.Date.max().date())
v['coupon_filled'] = int((df.CouponCode=='No Coupon').sum())
log.update(v)

summary = pd.DataFrame([(k,str(x)) for k,x in log.items()], columns=['Check','Result'])
with pd.ExcelWriter(OUT, engine='openpyxl', datetime_format='YYYY-MM-DD', date_format='YYYY-MM-DD') as w:
    df.to_excel(w,'Cleaned_Data',index=False)
    summary.to_excel(w,'Verification',index=False)
    ws = w.sheets['Cleaned_Data']; ws.freeze_panes='A2'
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = max(len(str(col[0].value)),12)+3
    for r in range(2,ws.max_row+1):
        ws.cell(r,2).number_format='YYYY-MM-DD'
        ws.cell(r,6).number_format='0.00'; ws.cell(r,14).number_format='0.00'
    w.sheets['Verification'].column_dimensions['A'].width=28; w.sheets['Verification'].column_dimensions['B'].width=60
df.to_csv(OUT.replace('.xlsx','.csv'),index=False,date_format='%Y-%m-%d')
json.dump(log,open(OUT.replace('.xlsx','_log.json'),'w'),indent=1)
print(json.dumps(log,indent=1))
