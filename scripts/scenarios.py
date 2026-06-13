import json, random
random.seed(7)
d=json.load(open("combos_data.json"))
picks=d["picks"]; odds={p[0]:p[3] for p in picks}; tier={p[0]:p[4] for p in picks}
combos=[c["legs"] for c in d["combos"]]
ids=[p[0] for p in picks]; STAKE=20; TOTAL=len(combos)*STAKE
def co(c):
    p=1.0
    for i in c:p*=odds[i]
    return p
def sim(prob, trials=200000):
    anyw=0; tot=0.0; cw=[0]*len(combos)
    for _ in range(trials):
        hit={i:(random.random()<prob[i]) for i in ids}
        ret=0.0; w=False
        for k,c in enumerate(combos):
            if all(hit[i] for i in c):
                ret+=co(c)*STAKE; w=True; cw[k]+=1
        if w:anyw+=1
        tot+=ret
    return anyw/trials*100, tot/trials, [round(x/trials*100,1) for x in cw]
scenarios={
 "Optimista (tus niveles: AA92 A85 M75)":{i:{"AA":.92,"A":.85,"M":.75}[tier[i]] for i in ids},
 "Mercado (prob = 1/cuota, lo que cree la casa)":{i:1/odds[i] for i in ids},
 "Realista (promedio de los dos)":{i:0.5*({"AA":.92,"A":.85,"M":.75}[tier[i]])+0.5*(1/odds[i]) for i in ids},
}
print(f"{'Escenario':48s} {'P(alguna gana)':>15s} {'Retorno prom':>13s} {'Neto prom':>11s}")
res={}
for name,pr in scenarios.items():
    aw,ar,cw=sim(pr)
    res[name]={"any":round(aw,1),"ret":round(ar,2),"net":round(ar-TOTAL,2)}
    print(f"{name:48s} {aw:>13.1f}% {ar:>12.0f} {ar-TOTAL:>+11.0f}")
print("\nApuesta total por ronda: S/",TOTAL)
json.dump(res,open("scenarios.json","w"),ensure_ascii=False,indent=1)
