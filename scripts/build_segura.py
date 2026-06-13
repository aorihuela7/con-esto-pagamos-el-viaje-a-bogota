import itertools, math, json
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------- 20 picks elegibles (ganador directo, cuota>=1.40) ----------
# id:(equipo, partido, cuota, miP, kickoff)
P={
1:("Escocia","Haiti vs Escocia",1.65,0.65,datetime(2026,6,13,20,0)),
2:("Turquia","Australia vs Turquia",1.72,0.80,datetime(2026,6,13,23,0)),
3:("Belgica","Belgica vs Egipto",1.68,0.60,datetime(2026,6,15,14,0)),
4:("Uruguay","Arabia Saudita vs Uruguay",1.47,0.70,datetime(2026,6,15,17,0)),
5:("Iran","Iran vs Nueva Zelanda",1.85,0.75,datetime(2026,6,15,20,0)),
6:("Argentina","Argentina vs Argelia",1.42,0.80,datetime(2026,6,16,20,0)),
7:("Colombia","Uzbekistan vs Colombia",1.41,0.90,datetime(2026,6,17,21,0)),
8:("Rep. Checa","Rep. Checa vs Sudafrica",1.67,0.75,datetime(2026,6,18,11,0)),
9:("Suiza","Suiza vs Bosnia",1.62,0.73,datetime(2026,6,18,14,0)),
10:("Marruecos","Escocia vs Marruecos",1.98,0.85,datetime(2026,6,19,17,0)),
11:("Belgica","Belgica vs Iran",1.44,0.77,datetime(2026,6,21,14,0)),
12:("Uruguay","Uruguay vs Cabo Verde",1.46,0.80,datetime(2026,6,21,17,0)),
13:("Egipto","Nueva Zelanda vs Egipto",1.80,0.75,datetime(2026,6,21,20,0)),
14:("Argentina","Argentina vs Austria",1.65,0.60,datetime(2026,6,22,12,0)),
15:("Argelia","Jordania vs Argelia",1.55,0.77,datetime(2026,6,22,22,0)),
16:("Croacia","Panama vs Croacia",1.55,0.85,datetime(2026,6,23,18,0)),
17:("Colombia","Colombia vs RD Congo",1.49,0.73,datetime(2026,6,23,21,0)),
18:("Brasil","Escocia vs Brasil",1.48,0.88,datetime(2026,6,24,17,0)),
19:("Corea del Sur","Sudafrica vs Corea del Sur",1.72,0.75,datetime(2026,6,24,20,0)),
20:("Senegal","Senegal vs Irak",1.45,0.75,datetime(2026,6,26,14,0)),
21:("Paises Bajos","Tunez vs Paises Bajos",1.55,0.83,datetime(2026,6,25,18,0)),
}
ids=list(P)
odds={i:P[i][2] for i in ids}
prob={i:P[i][3] for i in ids}
end ={i:P[i][4]+timedelta(hours=2) for i in ids}   # fin ~ inicio + 2h
def window(dt): return (dt-timedelta(hours=17)).date()   # ventana 17:00->16:59
STAKE=60.0; CAP_REFUND=60.0

def combo_window(c):  # ventana = la del partido que termina mas tarde
    last=max(c, key=lambda i:end[i]); return window(end[last]), last

def stats(c):
    o=math.prod(odds[i] for i in c)
    p_all=math.prod(prob[i] for i in c)
    # exactamente 1 falla
    p1=sum((1-prob[i])*math.prod(prob[j] for j in c if j!=i) for i in c)
    p2plus=1-p_all-p1
    payout=STAKE*o
    profit_all=payout-STAKE
    refund=min(STAKE,CAP_REFUND)           # reembolso si falla 1 -> neto 0
    ev=p_all*profit_all + p1*0 + p2plus*(-STAKE)
    return dict(odds=o,p_all=p_all,p1=p1,p2plus=p2plus,payout=payout,
                profit=profit_all,ev=ev)

# ---------- enumerar TODOS los combos de 5 patas ----------
cands=[]
for c in itertools.combinations(ids,5):
    o=math.prod(odds[i] for i in c)
    if o<5.30: continue
    st=stats(c); win,last=combo_window(c)
    cands.append((frozenset(c),st,win,last))
print("candidatos 5 patas (cuota>=5.30):",len(cands))

# ---------- seleccion: 12 combos, cap uso por pick, <=3 por ventana ----------
N=12; PICK_CAP=4; WIN_CAP=3
cands.sort(key=lambda x:-x[1]["ev"])   # mejor EV primero

def jaccard(a,b): return len(a&b)/len(a|b)

def select(pick_cap):
    chosen=[]; use={i:0 for i in ids}; wcount={}
    for c,st,win,last in cands:
        if len(chosen)>=N: break
        if any(use[i]>=pick_cap for i in c): continue
        if wcount.get(win,0)>=WIN_CAP: continue
        if any(jaccard(c,oc)>=0.8 for oc,_,_,_ in chosen): continue  # evita casi-duplicados
        chosen.append((c,st,win,last))
        for i in c: use[i]+=1
        wcount[win]=wcount.get(win,0)+1
    return chosen,use,wcount

chosen,use,wcount=select(PICK_CAP)
if len(chosen)<N:                      # relaja cap si no llega a 12
    chosen,use,wcount=select(5)
print("elegidos:",len(chosen))

# ---------- min hitting set (cuantos partidos deben fallar para tumbar TODAS) ----------
parlays=[c for c,_,_,_ in chosen]
def min_hit(maxk=5):
    uni=sorted(set().union(*parlays))
    for k in range(1,maxk+1):
        for combo in itertools.combinations(uni,k):
            cs=set(combo)
            if all(cs & set(p) for p in parlays): return k,combo
    return maxk+1,None
mh,mhset=min_hit()
print("min hitting set =",mh,"->",[P[i][0] for i in mhset] if mhset else None)

def cover_frac(k):
    uni=sorted(set().union(*parlays)); tot=cov=0
    for combo in itertools.combinations(uni,k):
        tot+=1; cs=set(combo)
        if any(not(cs&set(p)) for p in parlays): cov+=1
    return cov,tot
cov1=cover_frac(1); cov2=cover_frac(2)
print(f"cobertura k=1: {cov1[0]}/{cov1[1]}  k=2: {cov2[0]}/{cov2[1]}")
print("uso por pick:",{P[i][0]:use[i] for i in ids if use[i]})
print("combos por ventana:",{str(k):v for k,v in sorted(wcount.items())})

# ordena combos por ventana(fecha) y luego por EV
chosen.sort(key=lambda x:(x[2], -x[1]["ev"]))
print("\n# legs cuota   paga  P5gan P1fal P2+  EV    ventana")
out=[]
for k,(c,st,win,last) in enumerate(chosen,1):
    legs=sorted(c, key=lambda i:end[i])
    out.append((legs,st,win,last))
    print(f"{k:2d} {len(c)} {st['odds']:6.2f} S/{st['payout']:6.0f} "
          f"{st['p_all']*100:4.0f}% {st['p1']*100:4.0f}% {st['p2plus']*100:4.0f}% "
          f"S/{st['ev']:5.0f}  {win}  ({P[last][0]} {end[last].strftime('%d/%m %H:%M')})")

agg_ev=sum(st['ev'] for _,st,_,_ in chosen)
agg_cost=STAKE*len(chosen)
print(f"\ncosto total S/{agg_cost:.0f}  EV total(con seguro) S/{agg_ev:.0f}")

json.dump({
 "picks":{str(i):[P[i][0],P[i][1],odds[i],prob[i],P[i][4].strftime('%d/%m %H:%M')] for i in ids},
 "combos":[{"legs":sorted(c,key=lambda i:end[i]),"odds":round(st['odds'],2),
            "payout":round(st['payout']),"profit":round(st['profit']),
            "p_all":st['p_all'],"p1":st['p1'],"p2plus":st['p2plus'],"ev":st['ev'],
            "win":str(win),"last":P[last][0],"last_end":end[last].strftime('%d/%m %H:%M')}
           for c,st,win,last in chosen],
 "use":{str(i):use[i] for i in ids},
 "wcount":{str(k):v for k,v in sorted(wcount.items())},
 "mh":mh,"mhset":[P[i][0] for i in mhset] if mhset else [],
 "cov1":cov1,"cov2":cov2,"cost":agg_cost,"agg_ev":agg_ev,"stake":STAKE},
 open("segura.json","w"))
print("saved segura.json")
