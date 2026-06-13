import random, itertools, math, json
import numpy as np
random.seed(7); np.random.seed(7)

# 41 picks: id -> (label, partido, odds, miP, fecha)
P={
1:("Mexico gana","Mexico vs Sudafrica",1.45,0.75,"11/06"),
2:("Suiza gana","Catar vs Suiza",1.26,0.87,"13/06"),
3:("Escocia gana","Haiti vs Escocia",1.47,0.65,"17/06"),
4:("Turquia gana","Australia vs Turquia",1.80,0.78,"13/06"),
5:("Alemania gana","Alemania vs Curazao",1.05,0.98,"14/06"),
6:("Espana gana","Espana vs Cabo Verde",1.10,0.98,"15/06"),
7:("Belgica 1X","Belgica vs Egipto",1.19,0.77,"18/06"),
8:("Uruguay gana","Arabia Saudita vs Uruguay",1.50,0.6667,"15/06"),
9:("Francia gana","Francia vs Senegal",1.47,0.71,"16/06"),
10:("Noruega gana","Irak vs Noruega",1.23,0.87,"16/06"),
11:("Argentina gana","Argentina vs Argelia",1.42,0.68,"18/06"),
12:("Austria gana","Austria vs Jordania",1.34,0.82,"16/06"),
13:("Portugal gana","Portugal vs RD Congo",1.28,0.85,"17/06"),
14:("Ghana 1X","Ghana vs Panama",1.29,0.80,"17/06"),
15:("Colombia gana","Uzbekistan vs Colombia",1.42,0.92,"17/06"),
16:("Rep. Checa 1X","Rep. Checa vs Sudafrica",1.27,0.88,"18/06"),
17:("Suiza 1X","Suiza vs Bosnia",1.17,0.76,"18/06"),
18:("EE.UU. 1X","EE.UU. vs Australia",1.22,0.68,"19/06"),
19:("Marruecos gana","Escocia vs Marruecos",2.05,0.82,"19/06"),
20:("Brasil gana","Brasil vs Haiti",1.06,0.99,"19/06"),
21:("Ecuador gana","Ecuador vs Curazao",1.27,0.95,"20/06"),
22:("Japon 2X","Tunez vs Japon",1.20,0.8333,"20/06"),
23:("Espana gana","Espana vs Arabia Saudita",1.13,0.885,"21/06"),
24:("Belgica gana","Belgica vs Iran",1.47,0.73,"21/06"),
25:("Uruguay gana","Uruguay vs Cabo Verde",1.47,0.76,"21/06"),
26:("Egipto 2X","Nueva Zelanda vs Egipto",1.19,0.87,"21/06"),
27:("Argentina gana","Argentina vs Austria",1.67,0.55,"22/06"),
28:("Francia gana","Francia vs Irak",1.12,0.96,"22/06"),
29:("Argelia 2X","Jordania vs Argelia",1.18,0.71,"25/06"),
30:("Portugal gana","Portugal vs Uzbekistan",1.27,0.95,"23/06"),
31:("Inglaterra gana","Inglaterra vs Ghana",1.32,0.82,"23/06"),
32:("Croacia gana","Panama vs Croacia",1.50,0.84,"23/06"),
33:("Colombia gana","Colombia vs RD Congo",1.52,0.75,"23/06"),
34:("Brasil gana","Escocia vs Brasil",1.40,0.92,"24/06"),
35:("Marruecos gana","Marruecos vs Haiti",1.26,0.93,"24/06"),
36:("Corea 2X","Sudafrica vs Corea del Sur",1.27,0.80,"24/06"),
37:("C. Marfil gana","Curazao vs Costa de Marfil",1.34,0.84,"25/06"),
38:("Senegal gana","Senegal vs Irak",1.47,0.73,"26/06"),
39:("Belgica gana","Nueva Zelanda vs Belgica",1.30,0.86,"26/06"),
40:("Inglaterra gana","Panama vs Inglaterra",1.20,0.95,"27/06"),
41:("Argentina gana","Jordania vs Argentina",1.20,0.91,"27/06"),
}
NOVALUE={22,11,3,27,7,17,29,18}
ids=[i for i in P if i not in NOVALUE]           # 33 value picks
odds={i:P[i][2] for i in ids}
probs={i:P[i][3] for i in ids}
T=170.0   # combined-odds target -> profit at S/30 = 30*170-30 = S/5070
STAKE=30

# ---- candidate generation: ban some top picks each time to force diversity ----
TOP=sorted(ids,key=lambda x:-odds[x])[:11]
def gen():
    ban=set(random.sample(TOP, random.choice([0,1,1,2,2,3,3,4])))
    pool=[i for i in ids if i not in ban]
    order=sorted(pool, key=lambda x: odds[x]*(0.55+0.9*random.random()), reverse=True)
    chosen=[];prod=1.0
    for i in order:
        chosen.append(i);prod*=odds[i]
        if prod>=T: return frozenset(chosen)
        if len(chosen)>=17: break
    return None

cands=set()
for _ in range(60000):
    c=gen()
    if c: cands.add(c)
cands=[c for c in cands if 12<=len(c)<=17]
random.shuffle(cands)
cands=cands[:420]              # random mix of short (high-odds) and long (diverse) parlays
print("candidatos:",len(cands),"| reparto largos:",{n:sum(len(c)==n for c in cands) for n in range(12,18)})

# ---- triple coverage objective ----
allid=ids
idx={t:k for k,t in enumerate(itertools.combinations(sorted(allid),3))}
NT=len(idx)
def covered_triples(c):  # triples avoided by parlay c = triples within complement
    comp=sorted(set(allid)-set(c))
    return [idx[t] for t in itertools.combinations(comp,3)]
covset={i:set(covered_triples(c)) for i,c in enumerate(cands)}

def usage_of(chosen, skip=None):
    u={i:0 for i in ids}
    for j,ci in enumerate(chosen):
        if j==skip: continue
        for m in cands[ci]: u[m]+=1
    return u

def select(N, cap):
    chosen=[]; covd=set()
    for _ in range(N):
        u=usage_of(chosen)
        feas=[i for i in range(len(cands)) if i not in chosen and all(u[m]<cap for m in cands[i])]
        if feas:
            best=max(feas, key=lambda i: len(covset[i]-covd))
        else:  # no candidate respects cap: pick the one whose picks are least overused
            rem=[i for i in range(len(cands)) if i not in chosen]
            best=min(rem, key=lambda i: (max(u[m] for m in cands[i]), -len(covset[i]-covd)))
        chosen.append(best); covd|=covset[best]
    for sweep in range(12):
        improved=False
        for pos in range(len(chosen)):
            u=usage_of(chosen, skip=pos)
            others=set().union(*[covset[ci] for j,ci in enumerate(chosen) if j!=pos]) if len(chosen)>1 else set()
            base=len(others|covset[chosen[pos]])
            feas=[i for i in range(len(cands)) if i not in chosen and all(u[m]<cap for m in cands[i])]
            if not feas: continue
            cand=max(feas, key=lambda i: len(others|covset[i]))
            if len(others|covset[cand])>base:
                chosen[pos]=cand; improved=True
        if not improved: break
    return chosen

def min_hitting_set_le(parlays, maxk=4):
    # smallest set of matches that intersects EVERY parlay; return size if <=maxk else maxk+1
    universe=sorted(set().union(*[set(p) for p in parlays]))
    for k in range(1,maxk+1):
        for combo in itertools.combinations(universe,k):
            cs=set(combo)
            if all(cs & set(p) for p in parlays):
                return k, combo
    return maxk+1, None

best=None; bestscore=None
for N in (12,15,18):
    for cap in sorted(set([max(3,round(N*0.5)),max(3,round(N*0.6)),max(3,round(N*0.7)),N-2,N-1])):   # each pick in <=cap parlays
        sel=select(N,cap)
        parlays=[cands[i] for i in sel]
        if len(set(parlays))<N: continue
        mh,_=min_hitting_set_le(parlays,3)
        cur=set().union(*[covset[i] for i in sel]); unc=NT-len(cur)
        score=(mh, -unc, N)        # maximize min-hitting-set, then coverage, then N
        if bestscore is None or score>bestscore:
            bestscore=score; best=(N,cap,parlays,mh,unc)
N,cap,parlays,mh0,unc=best
print(f"MEJOR: N={N} cap={cap} min_hitting_set={mh0} triples_sin_cubrir={unc}/{NT}")
# order parlays by combined odds desc
parlays=sorted(parlays, key=lambda c: -math.prod(odds[i] for i in c))
print("\nELEGIDAS:",N)
data=[]
for k,c in enumerate(parlays,1):
    legs=sorted(c, key=lambda i:P[i][4])
    comb=math.prod(odds[i] for i in legs)
    pr=math.prod(probs[i] for i in legs)
    payout=STAKE*comb
    data.append({"legs":legs,"odds":round(comb,1),"prob":pr,"payout":payout,"n":len(legs)})
    print(f"#{k:2d} legs={len(legs):2d} cuota={comb:8.0f} pago=S/{payout:9.0f} probMia={pr*100:5.2f}%")

# ---- coverage stats ----
def cover_frac(parlays,k):
    U=sorted(set().union(*[set(p) for p in parlays]))
    tot=0;cov=0
    for combo in itertools.combinations(U,k):
        tot+=1
        cs=set(combo)
        if any(not(cs&set(p)) for p in parlays): cov+=1
    return cov,tot
for k in (1,2,3):
    cov,tot=cover_frac(parlays,k)
    print(f"cobertura k={k}: {cov}/{tot} = {100*cov/tot:.1f}%  (combinadas que sobreviven si fallan esos {k})")
mh,combo=min_hitting_set_le(parlays,5)
print("min hitting set =",mh, "-> se necesitan al menos",mh,"partidos fallando para tumbar TODAS")

# ---- Monte Carlo with miP ----
U=sorted(set().union(*[set(p) for p in parlays]))
pi={i:probs[i] for i in U}
SIM=300000
arr=np.array([pi[i] for i in U])
draws=(np.random.rand(SIM,len(U))<arr)  # True=gana
col={m:j for j,m in enumerate(U)}
hits=np.zeros(SIM,dtype=int)
paywin=np.zeros(SIM)
for d in data:
    mask=np.ones(SIM,dtype=bool)
    for m in d["legs"]: mask&=draws[:,col[m]]
    hits+=mask
    paywin+=mask*d["payout"]
cost=STAKE*len(parlays)
net=paywin-cost
p_any=(hits>=1).mean()*100
print(f"\nMonte Carlo {SIM:,} sims  (costo total S/{cost})")
print(f"P(al menos 1 combinada acierta) = {p_any:.1f}%")
print(f"E[# combinadas que aciertan]    = {hits.mean():.2f}")
print(f"E[pago bruto]   = S/{paywin.mean():.0f}")
print(f"E[neto]         = S/{net.mean():.0f}")
print(f"P(terminar ganando, neto>0) = {(net>0).mean()*100:.1f}%")

json.dump({"data":[{**d} for d in data],"P":{str(i):P[i] for i in ids},
           "cover":{str(k):cover_frac(parlays,k) for k in (1,2,3)},
           "minhit":mh,"cost":cost,"p_any":p_any,"e_hits":float(hits.mean()),
           "e_gross":float(paywin.mean()),"e_net":float(net.mean()),
           "p_win":float((net>0).mean()*100),"stake":STAKE,"sims":SIM,"N":len(parlays)},
          open("parlays.json","w"))
print("saved parlays.json")
