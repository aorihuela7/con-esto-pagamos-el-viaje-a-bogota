import json
import numpy as np
# value picks: id:(pick,odds,prob,date)
V={
1:("México gana",1.45,0.75,"11/06"),2:("Suiza gana",1.26,0.87,"13/06"),4:("Turquía gana",1.80,0.78,"13/06"),
5:("Alemania gana",1.05,0.98,"14/06"),6:("España gana",1.10,0.98,"15/06"),9:("Francia gana",1.47,0.71,"16/06"),
10:("Noruega gana",1.23,0.87,"16/06"),12:("Austria gana",1.34,0.82,"16/06"),13:("Portugal gana",1.28,0.85,"17/06"),
14:("Ghana 1X",1.29,0.80,"17/06"),15:("Colombia gana",1.42,0.92,"17/06"),16:("Rep.Checa 1X",1.27,0.88,"18/06"),
19:("Marruecos gana",2.05,0.82,"19/06"),20:("Brasil gana",1.06,0.99,"19/06"),21:("Ecuador gana",1.27,0.95,"20/06"),
24:("Bélgica gana",1.47,0.73,"21/06"),25:("Uruguay gana",1.47,0.76,"21/06"),26:("Egipto 2X",1.19,0.87,"21/06"),
28:("Francia gana",1.12,0.96,"22/06"),30:("Portugal gana",1.27,0.95,"23/06"),31:("Inglaterra gana",1.32,0.82,"23/06"),
32:("Croacia gana",1.50,0.84,"23/06"),33:("Colombia gana",1.52,0.75,"23/06"),34:("Brasil gana",1.40,0.92,"24/06"),
35:("Marruecos gana",1.26,0.93,"24/06"),36:("Corea 2X",1.27,0.80,"24/06"),37:("C.Marfil gana",1.34,0.84,"25/06"),
38:("Senegal gana",1.47,0.73,"26/06"),39:("Bélgica gana",1.30,0.86,"26/06"),40:("Inglaterra gana",1.20,0.95,"27/06"),
41:("Argentina gana",1.20,0.91,"27/06"),23:("España gana",1.13,0.885,"21/06"),
}
def dkey(d): dd,mm=d.split("/"); return (int(mm),int(dd))
edge={i:V[i][2]*V[i][1]-1 for i in V}
# top 30 by edge, snake draft into 6 routes
order=sorted(V,key=lambda i:-edge[i])[:30]
N=6
routes=[[] for _ in range(N)]
d=0; step=1
for idx,i in enumerate(order):
    routes[d].append(i)
    if d==N-1 and step==1: step=-1
    elif d==0 and step==-1: step=1
    else: d+=step; continue
    d+=0  # snake handled
# simpler snake
routes=[[] for _ in range(N)]
seq=list(range(N))+list(range(N-1,-1,-1))
k=0
for i in order:
    routes[seq[k%len(seq)]].append(i); k+=1
# sort each route by date
for r in routes: r.sort(key=lambda i:dkey(V[i][3]))

S0=20.0
out=[]
for ri,r in enumerate(routes,1):
    legs=[]; mult=1.0; surv=1.0
    for i in r:
        pick,odds,p,date=V[i]; mult*=odds; surv*=p
        legs.append({"id":i,"pick":pick,"date":date,"odds":odds,"prob":p,
                     "mult":round(mult,3),"surv":round(surv,4),
                     "stop_payout":round(S0*mult,2),"ev":round(surv*S0*mult,2)})
    # recommended exit: last leg where surv>=0.5
    rec=0
    for k,l in enumerate(legs):
        if l["surv"]>=0.50: rec=k
    out.append({"route":ri,"stake":S0,"legs":legs,"rec_exit_idx":rec})

# aggregate sim
np.random.seed(5)
def sim(use_market,trials=60000):
    # strategy: each route ride to its recommended exit (stop after rec leg); payout if all legs up to rec win
    ids=set(); 
    for r in out:
        for l in r["legs"]: ids.add(l["id"])
    ids=list(ids); idx={i:j for j,i in enumerate(ids)}
    P=np.array([ (1/V[i][1]) if use_market else V[i][2] for i in ids])
    M=np.random.random((trials,len(ids)))<P
    tot=np.zeros(trials); anyhit=np.zeros(trials,dtype=bool); staked=0.0
    for r in out:
        rec=r["rec_exit_idx"]; legs=r["legs"][:rec+1]
        staked+=r["stake"]
        cols=[idx[l["id"]] for l in legs]
        win=M[:,cols].all(axis=1)
        payout=r["stake"]*legs[-1]["mult"]
        tot+=win*payout; anyhit|=win
    net=tot-staked
    return {"staked":staked,"avg_ret":float(tot.mean()),"avg_net":float(net.mean()),
            "p_profit":float((net>0).mean()*100),"p_any":float(anyhit.mean()*100),
            "p10":float(np.percentile(net,10)),"p90":float(np.percentile(net,90))}

agg={"mine":sim(False),"market":sim(True)}
json.dump({"routes":out,"agg":agg,"S0":S0},open("routes.json","w"),ensure_ascii=False)

for r in out:
    print(f"\n=== RUTA {r['route']}  (stake S/{r['stake']:.0f}, salida sugerida tras leg {r['rec_exit_idx']+1}) ===")
    print(f"{'leg':3} {'fecha':6} {'pick':16} {'cuota':>5} {'miP':>4} {'mult':>6} {'vivo%':>6} {'cierra=S/':>9}")
    for k,l in enumerate(r['legs'],1):
        star=" <= salida" if k-1==r['rec_exit_idx'] else ""
        print(f"{k:3} {l['date']:6} {l['pick'][:16]:16} {l['odds']:>5.2f} {l['prob']*100:>3.0f}% {l['mult']:>6.2f} {l['surv']*100:>5.0f}% {l['stop_payout']:>9.2f}{star}")
print("\n=== AGREGADO (ride hasta salida sugerida) ===")
for k,a in agg.items():
    lbl="TUS PROB" if k=="mine" else "MERCADO"
    print(f"{lbl}: apostado S/{a['staked']:.0f} | retorno prom S/{a['avg_ret']:.0f} | neto prom S/{a['avg_net']:+.0f} | P(alguna ruta cobra) {a['p_any']:.0f}% | P(ronda rentable) {a['p_profit']:.0f}%")
