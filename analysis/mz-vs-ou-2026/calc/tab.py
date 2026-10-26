import json
D=json.load(open('data.json')); mz={int(k):v for k,v in D['mz'].items()}; ou={int(k):v for k,v in D['ou'].items()}
M=range(1,10)
def row(a,b): return (a,b,a-b,(a-b)/b*100 if b else None)
res={}
for m in M:
    z=mz[m];o=ou[m]
    dg=o['gross'][0]/o['gross'][1]
    res.setdefault('gross_m3',[]).append(row(z['gross_m3'],o['gross'][1]))
    res.setdefault('gross_t',[]).append(row(z['gross_t'] or z['gross_m3']*dg,o['gross'][0]))
    tm=z['tab_m3']+z['vrez_m3']; dt=o['tab'][0]/o['tab'][1]
    res.setdefault('tab_m3',[]).append(row(tm,o['tab'][1]))
    res.setdefault('tab_t',[]).append(row((z['tab_t']+z['vrez_t']) if z['tab_t'] else tm*dt,o['tab'][0]))
    res.setdefault('kv_m3',[]).append(row(z['kv_m3'],o['kv'][1]))
    res.setdefault('kv_t',[]).append(row(z['kv_t'],o['kv'][0]))
    res.setdefault('tok_m3',[]).append(row(z['tokko0_m3'],o['tokko'][1]))
    res.setdefault('tok_t',[]).append(row(z['tokko0_t'],o['tokko'][0]))
    for i,s in [(1,'ST-1'),(2,'ST-2'),(3,'ST-3')]:
        res.setdefault(f'tk{i}_m3',[]).append(row(z[f'tokko{i}_m3'],o['k_'+s][1]))
        res.setdefault(f'tk{i}_t',[]).append(row(z[f'tokko{i}_t'],o['k_'+s][0]))
f=lambda x:'—' if x is None else f'{x:,.1f}'.replace(',',' ')
for k,v in res.items():
    A=sum(r[0] for r in v);B=sum(r[1] for r in v)
    print(k,' | '.join(f'{m}:{f(r[0])}/{f(r[1])}/{f(r[2])}/{f(r[3])}%' for m,r in zip(M,v)),'|| TOT',f(A),f(B),f(A-B),f((A-B)/B*100))
# stages Gross/Tab
for m in (7,8):
    z=mz[m];o=ou[m]
    for s,val in z['grst'].items():
        b=o['g_'+s]; print('G',m,s,f(val),f(b[1]),f(val-b[1]),f((val-b[1])/b[1]*100),'| t',f(val*b[0]/b[1]),f(b[0]),f(val*b[0]/b[1]-b[0]))
    for s,val in list(z['tabst'].items())+[('VREZ',z['vrez_m3'])]:
        b=o['t_'+s]; print('T',m,s,f(val),f(b[1]),f(val-b[1]),f((val-b[1])/b[1]*100),'| t',f(val*b[0]/b[1]),f(b[0]),f(val*b[0]/b[1]-b[0]))
for m in (6,9):
    z=mz[m];o=ou[m];b=o['t_VREZ'];vt=z['vrez_t'] or z['vrez_m3']*b[0]/b[1]
    print('V',m,f(z['vrez_m3']),f(b[1]),f((z['vrez_m3']-b[1])/b[1]*100),'t',f(vt),f(b[0]),f((vt-b[0])/b[0]*100))
# reference MZ-file OU vs export
for m in M: print('REF',m,f(mz[m]['gross_ouf_m3']),f(ou[m]['gross'][1]),f(mz[m]['tab_ouf_m3']+mz[m]['vrez_ouf_m3']),f(ou[m]['tab'][1]))
