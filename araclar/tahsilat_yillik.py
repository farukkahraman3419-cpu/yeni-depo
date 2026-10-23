exec(open('mkxlsx.py').read().replace("wb.save('out/Musteri_Listesi.xlsx')","pass"))
names=[c.get('company') or c.get('name') for c in C]; assert len(set(names))==len(names)
cname={c['id']:(c.get('company') or c.get('name')) for c in C}
onos={o['id']:o.get('no','') for o in O}
PP=sorted(P,key=lambda p:(p.get('date') or '',p['id']))
years=sorted({(p.get('date') or '')[:4] for p in PP if p.get('date')})
hdrfill=PatternFill('solid',fgColor='1F3A5F'); money='"$"#,##0.00;[Red]-"$"#,##0.00;"-"'
def head(ws,row,cols):
    for j,h in enumerate(cols,1):
        c=ws.cell(row,j,h); c.font=F(bold=True,color='FFFFFF'); c.fill=hdrfill; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
# detay
wd=wb.create_sheet('Tahsilat Detayı')
wd['A1']='Tahsilat Detayı'; wd['A1'].font=F(bold=True,size=14)
wd['A2']=f'{len(PP)} tahsilat, tarihe göre sıralı. Kaynak: Sipariş ve Tahsilat Defteri, 23.10.2026.'; wd['A2'].font=F(italic=True,size=9,color='595959')
dh=['Tarih','Yıl','Müşteri / Firma','Yöntem','Sipariş no','Tutar (USD)','Alınan TL','Kur','Not']; head(wd,4,dh)
for i,p in enumerate(PP):
    r=5+i; d=dt(p.get('date') or '')
    usd=float(p.get('amount') or 0) if cur(p)=='USD' else 0
    vals=[d,f'=YEAR(A{r})',cname.get(p['customerId'],'Silinmiş müşteri'),p.get('method',''),onos.get(p.get('orderId'),''),usd,p.get('tlAmount') or (float(p.get('amount') or 0) if cur(p)=='TRY' else None),p.get('rate'),p.get('note','')]
    for j,v in enumerate(vals,1):
        c=wd.cell(r,j,v); c.font=F()
    wd.cell(r,1).number_format='dd.mm.yyyy'; wd.cell(r,6).number_format=money; wd.cell(r,7).number_format='#,##0.00" ₺";;"-"'; wd.cell(r,8).number_format='0.00##'
dl=4+len(PP)
for j,w in enumerate([12,7,32,14,14,15,15,9,45],1): wd.column_dimensions[get_column_letter(j)].width=w
wd.freeze_panes='A5'; wd.auto_filter.ref=f'A4:I{dl}'
# yıllık
wy=wb.create_sheet('Tahsilat (Yıllık)',1)
wy['A1']='Tahsilatlar – Yıllara Göre (USD)'; wy['A1'].font=F(bold=True,size=14)
wy['A2']="Rakamlar 'Tahsilat Detayı' sayfasından formülle toplanıyor. TL alınan ödemeler alındığı günkü kurla dolara çevrilmiş tutarıyla yer alır."; wy['A2'].font=F(italic=True,size=9,color='595959')
# özet tablo
head(wy,4,['Yıl','Tahsilat adedi','Toplam tahsilat (USD)'])
for k,y in enumerate(years):
    r=5+k; wy.cell(r,1,int(y)).font=F(bold=True)
    wy.cell(r,2,f"=COUNTIFS('Tahsilat Detayı'!$B$5:$B${dl},A{r})").font=F()
    c=wy.cell(r,3,f"=SUMIFS('Tahsilat Detayı'!$F$5:$F${dl},'Tahsilat Detayı'!$B$5:$B${dl},A{r})"); c.font=F(); c.number_format=money
ts=5+len(years); wy.cell(ts,1,'TOPLAM').font=F(bold=True)
wy.cell(ts,2,f'=SUM(B5:B{ts-1})').font=F(bold=True)
c=wy.cell(ts,3,f'=SUM(C5:C{ts-1})'); c.font=F(bold=True); c.number_format=money
for j in (1,2,3): wy.cell(ts,j).fill=PatternFill('solid',fgColor='E8EEF5')
# müşteri x yıl
R0=ts+3
cols=['Müşteri / Firma']+[str(y) for y in years]+['Toplam']
head(wy,R0,cols)
order=sorted(names,key=lambda n:-sum(float(p.get('amount') or 0) for p in PP if cname.get(p['customerId'])==n and cur(p)=='USD'))
for i,n in enumerate(order):
    r=R0+1+i; wy.cell(r,1,n).font=F()
    for k,y in enumerate(years):
        col=get_column_letter(2+k)
        c=wy.cell(r,2+k,f"=SUMIFS('Tahsilat Detayı'!$F$5:$F${dl},'Tahsilat Detayı'!$C$5:$C${dl},$A{r},'Tahsilat Detayı'!$B$5:$B${dl},{col}${R0})")
        c.font=F(); c.number_format=money
    lc=get_column_letter(1+len(years))
    c=wy.cell(r,2+len(years),f'=SUM(B{r}:{lc}{r})'); c.font=F(bold=True); c.number_format=money
RL=R0+len(order); RT=RL+1
wy.cell(RT,1,'TOPLAM').font=F(bold=True)
for k in range(len(years)+1):
    col=get_column_letter(2+k); c=wy.cell(RT,2+k,f'=SUM({col}{R0+1}:{col}{RL})'); c.font=F(bold=True); c.number_format=money
for j in range(1,len(cols)+1): wy.cell(RT,j).fill=PatternFill('solid',fgColor='E8EEF5')
# başlıktaki yıllar sayı olmalı (SUMIFS eşleşmesi için)
for k,y in enumerate(years): wy.cell(R0,2+k).value=int(y)
wy.column_dimensions['A'].width=32
for j in range(2,len(cols)+1): wy.column_dimensions[get_column_letter(j)].width=15
wy.freeze_panes=wy.cell(R0+1,2); wy.auto_filter.ref=f'A{R0}:{get_column_letter(len(cols))}{RL}'
wb.save('out/Musteri_Listesi.xlsx'); print(years, len(PP))
