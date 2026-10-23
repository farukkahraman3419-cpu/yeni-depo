import openpyxl, sys, re, datetime, json, collections, hashlib
wb=openpyxl.load_workbook('xl/imalat.xlsx',data_only=True,read_only=True)
SKIP={'SOFANNY (2)','ALLAH BELANI VERİR ATARSAN','KUTULAMACI MUSTAFA','MATBACI MEHMET BEY TL','MATBACI MEHMET BEY USD','HANİFE HANIM MATBAA','VAKUMCU HÜSEYİN','VAKUMCU HÜSEYİN borcumuz var'}
def h(v): return re.sub(r'\s+',' ',str(v or '')).strip().upper()
def num(v):
    if v in (None,''): return 0.0
    return float(v) if isinstance(v,(int,float)) else None
def dt(v):
    if isinstance(v,datetime.datetime): return v.date().isoformat()
    return None
def slug(s): return 'x'+hashlib.sha1(s.encode()).hexdigest()[:10]
customers=[];orders=[];payments=[];report=[]
NOW='2026-10-22T22:30:00Z'
for ws in wb.worksheets:
    if ws.title in SKIP: continue
    rows=list(ws.iter_rows(values_only=True)); H=[h(v) for v in rows[0]]
    ix=lambda n: H.index(n) if n in H else None
    c=dict(name=ix('ÜRÜN CİNSİ'),code=ix('ÜRÜN KODU'),qty=ix('ADET'),tu=ix('TUTAR USD'),pu=ix('ALINAN ÖDEME USD'),ku=ix('KALAN USD'),tt=ix('TUTAR TL'),pt=ix('ALINAN ÖDEME TL'),note=ix('AÇIKLAMA'))
    title=re.sub(r'\s+',' ',ws.title).strip()
    cid=slug('c:'+ws.title)
    customers.append(dict(id=cid,company=title,name='',phone='',email='',city='',taxNo='',address='',note=f'Excel\'den aktarıldı: "{title}" sayfası.',createdAt=NOW,updatedAt=NOW,src='excel'))
    g=lambda r,k: r[c[k]] if c[k] is not None and c[k]<len(r) else None
    lastdate=None; firstdate=None; prevK=0.0; lastK=0.0
    cur_order=None; seq=0; adj=[]
    def flush():
        global cur_order
        if cur_order and cur_order['items']: orders.append(cur_order)
        cur_order=None
    def new_order(date, kind=None):
        global seq
        seq+=1
        return dict(id=slug(f'o:{ws.title}:{seq}'),customerId=cid,date=date,status='Teslim Edildi',vatRate=0,currency='USD',items=[],note='',createdAt=NOW,updatedAt=NOW,src='excel',**({'kind':kind} if kind else {}))
    tl_entries=[]
    for i,r in enumerate(rows[1:],2):
        r=list(r)
        if not any(v not in (None,'') for v in r): continue
        d=dt(r[0])
        if d: lastdate=d; firstdate=firstdate or d
        date=lastdate or firstdate or '2024-01-01'
        nm=re.sub(r'\s+',' ',str(g(r,'name') or '')).strip(); code=str(g(r,'code') or '').strip()
        qty=num(g(r,'qty')) or 0.0; tu=num(g(r,'tu')); pu=num(g(r,'pu')); tt=num(g(r,'tt')) or 0.0; pt=num(g(r,'pt')) or 0.0
        k=g(r,'ku'); note=g(r,'note'); note='' if note in (None,'') or (isinstance(note,str) and note.startswith('#')) else note
        if tu is None: tu=0.0
        if pu is None: pu=0.0
        has=bool(nm or tu or pu)
        if isinstance(k,(int,float)) and has:
            jump=k-(prevK+tu-pu)
            if abs(jump)>0.005:
                flush()
                o=new_order(date,'devir' if 'DEVR' in nm.upper() or i==2 else 'duzeltme')
                o['items']=[dict(productId='',name='Yıl devri (önceki bakiye)' if o.get('kind')=='devir' else 'Bakiye düzeltmesi (Excel)',sku='',qty=1,unitPrice=round(jump,2))]
                o['note']=f'Excel satır {i}'; orders.append(o); adj.append((i,round(jump,2)))
            prevK=k; lastK=k
        elif has and (tu or pu):
            prevK=prevK+tu-pu
        if pu:
            flush()
            rate=None
            if isinstance(note,(int,float)) and 1<note<200: rate=round(note,4)
            elif pt and pu: rate=round(pt/pu,4)
            pnote=[]
            if pt: pnote.append(f'{pt:,.2f} TL'.replace(',','X').replace('.',',').replace('X','.')+(f' · kur {rate}' if rate else ''))
            if isinstance(note,str) and note: pnote.append(note)
            payments.append(dict(id=slug(f'p:{ws.title}:{i}'),customerId=cid,orderId='',amount=round(pu,2),currency='USD',tlAmount=round(pt,2) if pt else None,rate=rate,date=date,method='Havale / EFT',note=' · '.join(pnote) or (nm if nm.upper() not in ('TAHSİLAT','TAHSILAT') else ''),createdAt=NOW,updatedAt=NOW,src='excel'))
        if tu:
            if not cur_order or cur_order['date']!=date: flush(); cur_order=new_order(date)
            q=qty if qty else 1
            cur_order['items'].append(dict(productId='',name=nm or 'Kalem',sku=code if code not in ('None',) else '',qty=q,unitPrice=round(tu/q,6)))
            if isinstance(note,str) and note and note not in cur_order['note']: cur_order['note']=(cur_order['note']+' · '+note).strip(' ·')
        if not tu and not pu and (tt or pt) and nm:
            tl_entries.append((i,date,nm,tt,pt))
    flush()
    # compute balance from docs
    bal=sum(sum(it['qty']*it['unitPrice'] for it in o['items']) for o in orders if o['customerId']==cid)-sum(p['amount'] for p in payments if p['customerId']==cid)
    report.append(dict(sheet=ws.title,excel=round(lastK,2),calc=round(bal,2),adj=adj,tl=tl_entries,n_o=sum(1 for o in orders if o['customerId']==cid),n_p=sum(1 for p in payments if p['customerId']==cid)))
# order numbers
orders.sort(key=lambda o:(o['date'],o['customerId'],o['id']))
cnt=collections.Counter()
for o in orders:
    y=o['date'][:4]; cnt[y]+=1; o['no']=f'SP-{y}-{cnt[y]:04d}'
json.dump(dict(customers=customers,orders=orders,payments=payments),open('import.json','w'),ensure_ascii=False,default=str)
json.dump(report,open('report.json','w'),ensure_ascii=False,default=str)
print('customers',len(customers),'orders',len(orders),'payments',len(payments),'items',sum(len(o['items']) for o in orders))
print('years',dict(cnt))
bad=[r for r in report if abs(r['excel']-r['calc'])>0.05]
print('mismatch',[(r['sheet'],r['excel'],r['calc']) for r in bad])
print('adjust (non-first-row):',[(r['sheet'],[a for a in r['adj'] if a[0]!=2]) for r in report if any(a[0]!=2 for a in r['adj'])])
print('TL rows:',sum(len(r['tl']) for r in report), [(r['sheet'],len(r['tl'])) for r in report if r['tl']])
print('max order doc bytes', max(len(json.dumps(o,ensure_ascii=False)) for o in orders))
