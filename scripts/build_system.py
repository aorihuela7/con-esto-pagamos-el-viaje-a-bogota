import json, random
import numpy as np
random.seed(11); np.random.seed(11)

picks = [
 (1,"México gana","México vs Sudáfrica",1.45),(2,"Suiza gana","Catar vs Suiza",1.26),
 (3,"Escocia gana","Haití vs Escocia",1.47),(4,"Turquía gana","Australia vs Turquía",1.80),
 (5,"Alemania gana","Alemania vs Curazao",1.05),(6,"España gana","España vs Cabo Verde",1.10),
 (7,"Bélgica o empate (1X)","Bélgica vs Egipto",1.19),(8,"Uruguay gana","Arabia Saudita vs Uruguay",1.50),
 (9,"Francia gana","Francia vs Senegal",1.47),(10,"Noruega gana","Irak vs Noruega",1.23),
 (11,"Argentina gana","Argentina vs Argelia",1.42),(12,"Austria gana","Austria vs Jordan",1.34),
 (13,"Portugal gana","Portugal vs RD Congo",1.28),(14,"Ghana o empate (1X)","Ghana vs Panamá",1.29),
 (15,"Colombia gana","Uzbekistán vs Colombia",1.42),(16,"Rep. Checa o empate (1X)","Rep. Checa vs Sudáfrica",1.27),
 (17,"Suiza o empate (1X)","Suiza vs Bosnia",1.17),(18,"EE.UU. o empate (1X)","EE.UU. vs Australia",1.22),
 (19,"Marruecos gana","Escocia vs Marruecos",2.05),(20,"Brasil gana","Brasil vs Haití",1.06),
 (21,"Ecuador gana","Ecuador vs Curazao",1.27),(22,"Japón o empate (2X)","Túnez vs Japón",1.20),
 (23,"España gana","España vs Arabia Saudita",1.13),(24,"Bélgica gana","Bélgica vs Irán",1.47),
 (25,"Uruguay gana","Uruguay vs Cabo Verde",1.47),(26,"Egipto o empate (2X)","Nueva Zelanda vs Egipto",1.19),
 (27,"Argentina gana","Argentina vs Austria",1.67),(28,"Francia gana","Francia vs Irak",1.12),
 (29,"Argelia o empate (2X)","Jordan vs Argelia",1.18),(30,"Portugal gana","Portugal vs Uzbekistán",1.27),
 (31,"Inglaterra gana","Inglaterra vs Ghana",1.32),(32,"Croacia gana","Panamá vs Croacia",1.50),
 (33,"Colombia gana","Colombia vs RD Congo",1.52),(34,"Brasil gana","Escocia vs Brasil",1.40),
 (35,"Marruecos gana","Marruecos vs Haití",1.26),(36,"Corea del Sur o empate (2X)","Sudáfrica vs Corea del Sur",1.27),
 (37,"Costa de Marfil gana","Curazao vs Costa de Marfil",1.34),(38,"Senegal gana","Senegal vs Irak",1.47),
 (39,"Bélgica gana","Nueva Zelanda vs Bélgica",1.30),(40,"Inglaterra gana","Panamá vs Inglaterra",1.20),
 (41,"Argentina gana","Jordan vs Argentina",1.20),
]
ids=[p[0] for p in picks]
odds={p[0]:p[3] for p in picks}
pm={p[0]:round(1/p[3],4) for p in picks}   # market implied prob (incluye margen)

# ---- Portfolio template (barbell) ----
# (name, n_legs, candidate_pool_size_by_prob_rank)
ranked=sorted(ids,key=lambda i:-pm[i])  # most probable first
def pool(n): return set(ranked[:n])
template=[]
for k in range(5): template.append((f"A{k+1}",5,pool(16)))   # anchors: high prob
for k in range(5): template.append((f"M{k+1}",10,pool(30)))  # medium
for k in range(4): template.append((f"L{k+1}",20,set(ids)))  # long lottery

# diversified per-band assignment (prob-weighted greedy, distinct per combo)
def assign_band(combos_n, legs_n, poolset):
    poolset=list(poolset)
    slots=combos_n*legs_n
    tot=sum(pm[i] for i in poolset)
    remaining={i: slots*pm[i]/tot for i in poolset}
    band=[]
    for c in range(combos_n):
        order=sorted(poolset,key=lambda i:(-remaining[i],random.random()))
        chosen=order[:legs_n]
        for i in chosen: remaining[i]-=1
        band.append(sorted(int(x) for x in chosen))
    return band

combos=[]
combos+=assign_band(5,5,pool(14))
combos+=assign_band(5,10,pool(30))
combos+=assign_band(4,20,set(ids))

def incl_count(cs):
    cnt={i:0 for i in ids}
    for c in cs:
        for i in c: cnt[i]+=1
    return cnt

# final coverage check
cnt=incl_count(combos)
excl={i:len(combos)-cnt[i] for i in ids}
assert all(excl[i]>=1 for i in ids), "a pick is in every combo"
# diversity check: no two combos identical
assert len(set(tuple(c) for c in combos))==len(combos), "duplicate combos"

def cuota(c): 
    r=1.0
    for i in c: r*=odds[i]
    return r
def prob(c):
    r=1.0
    for i in c: r*=pm[i]
    return r

# provisional stakes (placeholder, editable later)
stake_by_band={"A":30,"M":15,"L":5}
info=[]
for (name,n,pl),c in zip(template,combos):
    band=name[0]; st=stake_by_band[band]
    o=cuota(c); pr=prob(c)
    info.append({"name":name,"band":band,"legs":[int(x) for x in sorted(c)],"n":len(c),
                 "cuota":round(o,2),"prob":round(pr,4),"stake":st,
                 "payout":round(o*st,2),"ev":round(pr*o*st-st,2)})

# portfolio MC (market probs) -> realistic baseline
TR=200000
P=np.array([pm[i] for i in ids])
idx={i:j for j,i in enumerate(ids)}
M=np.random.random((TR,len(ids)))<P
combo_idx=[[idx[i] for i in ci["legs"]] for ci in info]
stakes=np.array([ci["stake"] for ci in info])
cuotas=np.array([ci["cuota"] for ci in info])
wins=np.zeros((TR,len(info)),dtype=bool)
for k,ci in enumerate(combo_idx):
    wins[:,k]=M[:,ci].all(axis=1)
ret=(wins*cuotas*stakes).sum(axis=1)
totstake=stakes.sum()
net=ret-totstake
mc={"any_win_pct":round(wins.any(axis=1).mean()*100,2),
    "avg_ret":round(ret.mean(),2),"total_stake":float(totstake),
    "avg_net":round(net.mean(),2),"p_profit":round((net>0).mean()*100,2),
    "per_combo_winrate":[round(wins[:,k].mean()*100,2) for k in range(len(info))]}

out={"picks":picks,"pm":pm,"combos":info,"excl":excl,"mc":mc}
json.dump(out,open("system_data.json","w"),ensure_ascii=False,indent=1)

print("PORTAFOLIO (cuotas de mercado, prob=1/cuota)")
print(f"{'Combo':5s} {'patas':5s} {'cuota':>9s} {'P(gana)':>8s} {'apuesta':>8s} {'pago':>10s} {'EV':>9s}")
for ci in info:
    print(f"{ci['name']:5s} {ci['n']:>5d} {ci['cuota']:>9.2f} {ci['prob']*100:>7.2f}% S/{ci['stake']:>5d} {ci['payout']:>10.2f} {ci['ev']:>+9.2f}")
print("\nBaseline con prob de mercado (200k sims):")
print(" P(alguna combinada gana): ",mc["any_win_pct"],"%")
print(" P(ronda rentable, net>0): ",mc["p_profit"],"%")
print(" Apuesta total: S/",mc["total_stake"]," | Retorno prom: S/",mc["avg_ret"]," | Neto prom: S/",mc["avg_net"])
