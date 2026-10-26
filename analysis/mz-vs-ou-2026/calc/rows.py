exec(open('tab.py').read().split('f=lambda')[0])
MN=['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь']
f=lambda x:f'{x:,.1f}'.replace(',',' ').replace('-','−')
def sg(x): return ('+' if x>0 else '')+f(x)
def pc(x):
    s=sg(x)+'%'
    return f'<b>{s}</b>' if abs(x)>5 else s
def td(x,c=''): return f'<td style="text-align:right{c}">{x}</td>'
for k in ['gross_m3','gross_t','kv_m3','kv_t','tab_m3','tab_t','tok_m3','tok_t']:
    print('=== ',k)
    v=res[k]
    for i,r in enumerate(v):
        col=';color:#B4561B' if abs(r[3])>5 else ''
        print(f'<tr><td>{MN[i]}</td>'+td(f(r[0]))+td(f(r[1]))+td(sg(r[2]))+td(pc(r[3]),col)+'</tr>')
    A=sum(r[0] for r in v);B=sum(r[1] for r in v)
    print(f'<tr style="background:#E7E4DC"><td><b>Итого 9 мес.</b></td>'+td('<b>'+f(A)+'</b>')+td('<b>'+f(B)+'</b>')+td('<b>'+sg(A-B)+'</b>')+td('<b>'+sg((A-B)/B*100)+'%</b>')+'</tr>')
    mx=max(range(9),key=lambda i:abs(v[i][3])); print('MAX',MN[mx],sg(v[mx][3]),'n>5:',sum(abs(r[3])>5 for r in v), 'mean abs', round(sum(abs(r[3]) for r in v)/9,1))
for u in ['m3','t']:
    print('=== tok st',u)
    for i in range(9):
        cells=''
        for s in (1,2,3):
            r=res[f'tk{s}_{u}'][i]; col=';color:#B4561B' if abs(r[3])>5 else ''
            cells+=td(f(r[0]))+td(f(r[1]))+td(pc(r[3]),col)
        print(f'<tr><td>{MN[i]}</td>{cells}</tr>')
    cells=''
    for s in (1,2,3):
        v=res[f'tk{s}_{u}'];A=sum(r[0] for r in v);B=sum(r[1] for r in v)
        cells+=td('<b>'+f(A)+'</b>')+td('<b>'+f(B)+'</b>')+td('<b>'+sg((A-B)/B*100)+'%</b>')
    print(f'<tr style="background:#E7E4DC"><td><b>Итого</b></td>{cells}</tr>')
