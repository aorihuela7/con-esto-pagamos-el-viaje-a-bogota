import json, random
random.seed(42)

# id, pick_label, match, odds, tier
picks = [
 (1,"México gana","México vs Sudáfrica",1.45,"A"),
 (2,"Suiza gana","Catar vs Suiza",1.26,"A"),
 (3,"Escocia gana","Haití vs Escocia",1.47,"M"),
 (4,"Turquía gana","Australia vs Turquía",1.80,"A"),
 (5,"Alemania gana","Alemania vs Curazao",1.05,"AA"),
 (6,"España gana","España vs Cabo Verde",1.10,"AA"),
 (7,"Bélgica o empate (1X)","Bélgica vs Egipto",1.19,"M"),
 (8,"Uruguay gana","Arabia Saudita vs Uruguay",1.50,"M"),
 (9,"Francia gana","Francia vs Senegal",1.47,"M"),
 (10,"Noruega gana","Irak vs Noruega",1.23,"AA"),
 (11,"Argentina gana","Argentina vs Argelia",1.42,"A"),
 (12,"Austria gana","Austria vs Jordan",1.34,"A"),
 (13,"Portugal gana","Portugal vs RD Congo",1.28,"AA"),
 (14,"Ghana o empate (1X)","Ghana vs Panamá",1.29,"A"),
 (15,"Colombia gana","Uzbekistán vs Colombia",1.42,"A"),
 (16,"Rep. Checa o empate (1X)","Rep. Checa vs Sudáfrica",1.27,"A"),
 (17,"Suiza o empate (1X)","Suiza vs Bosnia",1.17,"A"),
 (18,"EE.UU. o empate (1X)","EE.UU. vs Australia",1.22,"M"),
 (19,"Marruecos gana","Escocia vs Marruecos",2.05,"A"),
 (20,"Brasil gana","Brasil vs Haití",1.06,"AA"),
 (21,"Ecuador gana","Ecuador vs Curazao",1.27,"AA"),
 (22,"Japón o empate (2X)","Túnez vs Japón",1.20,"A"),
 (23,"España gana","España vs Arabia Saudita",1.13,"A"),
 (24,"Bélgica gana","Bélgica vs Irán",1.47,"A"),
 (25,"Uruguay gana","Uruguay vs Cabo Verde",1.47,"M"),
 (26,"Egipto o empate (2X)","Nueva Zelanda vs Egipto",1.19,"A"),
 (27,"Argentina gana","Argentina vs Austria",1.67,"M"),
 (28,"Francia gana","Francia vs Irak",1.12,"AA"),
 (29,"Argelia o empate (2X)","Jordan vs Argelia",1.18,"M"),
 (30,"Portugal gana","Portugal vs Uzbekistán",1.27,"AA"),
 (31,"Inglaterra gana","Inglaterra vs Ghana",1.32,"A"),
 (32,"Croacia gana","Panamá vs Croacia",1.50,"A"),
 (33,"Colombia gana","Colombia vs RD Congo",1.52,"A"),
 (34,"Brasil gana","Escocia vs Brasil",1.40,"AA"),
 (35,"Marruecos gana","Marruecos vs Haití",1.26,"AA"),
 (36,"Corea del Sur o empate (2X)","Sudáfrica vs Corea del Sur",1.27,"A"),
 (37,"Costa de Marfil gana","Curazao vs Costa de Marfil",1.34,"A"),
 (38,"Senegal gana","Senegal vs Irak",1.47,"M"),
 (39,"Bélgica gana","Nueva Zelanda vs Bélgica",1.30,"AA"),
 (40,"Inglaterra gana","Panamá vs Inglaterra",1.20,"AA"),
 (41,"Argentina gana","Jordan vs Argentina",1.20,"AA"),
]

N_COMBOS = 12
LEGS = 20
odds = {p[0]: p[3] for p in picks}
tier = {p[0]: p[4] for p in picks}
ids = [p[0] for p in picks]

# Target inclusion counts per pick (out of 12)
target = {}
for i in ids:
    if tier[i]=="AA": target[i]=9
    elif tier[i]=="A": target[i]=5
    else: target[i]=2  # M

# total slots needed = 12*20 = 240
need = N_COMBOS*LEGS
cur = sum(target.values())
# distribute remaining (need-cur) to A picks (bump some A from 5->6)
A_ids = [i for i in ids if tier[i]=="A"]
diff = need - cur
j=0
while diff>0:
    target[A_ids[j%len(A_ids)]] += 1
    j+=1; diff-=1
while diff<0:
    target[A_ids[j%len(A_ids)]] -= 1
    j+=1; diff+=1
assert sum(target.values())==need

# Greedy balanced assignment: for each combo pick the 20 picks most "owed"
remaining = dict(target)
combos = []
for c in range(N_COMBOS):
    # sort by remaining desc, tiebreak random for spread
    order = sorted(ids, key=lambda i:(-remaining[i], random.random()))
    chosen = order[:LEGS]
    for i in chosen:
        remaining[i]-=1
    combos.append(sorted(chosen))

# Verify
assert all(len(set(c))==LEGS for c in combos), "combo size error"
assert all(remaining[i]==0 for i in ids), "remaining not zero"

# Coverage: how many combos exclude each pick
excl = {i: sum(1 for c in combos if i not in c) for i in ids}
incl = {i: N_COMBOS-excl[i] for i in ids}
assert all(excl[i]>=1 for i in ids), "a pick is in every combo (no coverage)"

# Combo odds
def combo_odds(c): 
    p=1.0
    for i in c: p*=odds[i]
    return p
STAKE=20
combo_info=[]
for idx,c in enumerate(combos,1):
    o=combo_odds(c)
    combo_info.append({"combo":idx,"legs":c,"odds":round(o,2),"payout":round(o*STAKE,2),
                       "n_AA":sum(1 for i in c if tier[i]=="AA"),
                       "n_A":sum(1 for i in c if tier[i]=="A"),
                       "n_M":sum(1 for i in c if tier[i]=="M")})

# Monte Carlo
P={"AA":0.92,"A":0.85,"M":0.75}
TRIALS=200000
combo_wins=[0]*N_COMBOS
any_win=0
total_return=0.0
import random as r2
for t in range(TRIALS):
    hit={i:(random.random()<P[tier[i]]) for i in ids}
    won_any=False; ret=0.0
    for k,c in enumerate(combos):
        if all(hit[i] for i in c):
            combo_wins[k]+=1; won_any=True
            ret+=combo_odds(c)*STAKE
    if won_any: any_win+=1
    total_return+=ret
total_stake=N_COMBOS*STAKE
mc={
 "trials":TRIALS,
 "per_combo_winrate":[round(w/TRIALS*100,2) for w in combo_wins],
 "portfolio_any_win_pct":round(any_win/TRIALS*100,2),
 "avg_return_per_round":round(total_return/TRIALS,2),
 "total_stake":total_stake,
 "avg_net":round(total_return/TRIALS-total_stake,2),
 "expected_failures_of_41":round(sum(1-P[tier[i]] for i in ids),2),
}

out={"picks":picks,"target":target,"combos":combo_info,"coverage_incl":incl,"coverage_excl":excl,"mc":mc}
with open("/sessions/elegant-lucid-cray/mnt/outputs/combos_data.json","w") as f:
    json.dump(out,f,ensure_ascii=False,indent=1)

print("=== COMBOS (12 x 20 legs) ===")
for ci in combo_info:
    print(f"Combo {ci['combo']:2d}: cuota={ci['odds']:>9.2f}  pago(20)={ci['payout']:>11.2f}  [AA{ci['n_AA']} A{ci['n_A']} M{ci['n_M']}]")
print("\n=== COVERAGE (en cuántas combinadas aparece cada pick) ===")
for i in ids:
    print(f"  #{i:2d} {tier[i]:2s} -> en {incl[i]:2d}/12, fuera de {excl[i]:2d}")
print("\n=== MONTE CARLO (probabilidades supuestas AA92% A85% M75%) ===")
print("Trials:",mc["trials"])
print("Win rate por combo (%):",mc["per_combo_winrate"])
print("Prob. de que ALGUNA combinada gane: ",mc["portfolio_any_win_pct"],"%")
print("Fallos esperados de los 41 picks: ",mc["expected_failures_of_41"])
print("Retorno bruto promedio por ronda:  S/",mc["avg_return_per_round"])
print("Apuesta total:                     S/",mc["total_stake"])
print("Neto promedio por ronda:           S/",mc["avg_net"])
