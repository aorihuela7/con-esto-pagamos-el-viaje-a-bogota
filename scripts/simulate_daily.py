import numpy as np, json
np.random.seed(3)
# id: (pick, odds, myprob)
P={
1:("México gana",1.45,0.75),2:("Suiza gana",1.26,0.87),4:("Turquía gana",1.80,0.78),
5:("Alemania gana",1.05,0.98),6:("España gana",1.10,0.98),8:("Uruguay gana",1.50,0.6667),
9:("Francia gana",1.47,0.71),10:("Noruega gana",1.23,0.87),12:("Austria gana",1.34,0.82),
13:("Portugal gana",1.28,0.85),14:("Ghana 1X",1.29,0.80),15:("Colombia gana",1.42,0.92),
16:("Rep.Checa 1X",1.27,0.88),19:("Marruecos gana",2.05,0.82),20:("Brasil gana",1.06,0.99),
21:("Ecuador gana",1.27,0.95),23:("España gana",1.13,0.885),24:("Bélgica gana",1.47,0.73),
25:("Uruguay gana",1.47,0.76),26:("Egipto 2X",1.19,0.87),28:("Francia gana",1.12,0.96),
30:("Portugal gana",1.27,0.95),31:("Inglaterra gana",1.32,0.82),32:("Croacia gana",1.50,0.84),
33:("Colombia gana",1.52,0.75),34:("Brasil gana",1.40,0.92),35:("Marruecos gana",1.26,0.93),
36:("Corea 2X",1.27,0.80),37:("C.Marfil gana",1.34,0.84),38:("Senegal gana",1.47,0.73),
39:("Bélgica gana",1.30,0.86),40:("Inglaterra gana",1.20,0.95),41:("Argentina gana",1.20,0.91),
}
days=[("11/06",[1]),("13/06",[2,4]),("14/06",[5]),("15/06",[6,8]),("16/06",[9,10,12]),
("17/06",[13,14,15]),("18/06",[16]),("19/06",[19,20]),("20/06",[21]),("21/06",[23,24,25,26]),
("22/06",[28]),("23/06",[30,31,32,33]),("24/06",[34,35,36]),("25/06",[37]),("26/06",[38,39]),
("27/06",[40,41])]

KFRAC=0.65; CAP_SINGLE=0.06; CAP_DAY=0.25; COMBO_FRAC=0.02; LOCK=0.40
def kelly(p,b):
    f=(p*b-1)/(b-1)
    return max(0.0,f)
# fixed stake fractions per pick (of day-start bank)
sfrac={i: min(CAP_SINGLE, KFRAC*kelly(P[i][2],P[i][1])) for i in P}

def simulate(trials, use_market):
    start=500.0
    finals=np.zeros(trials)
    minbank=np.zeros(trials)
    for t in range(trials):
        bank=start; locked=0.0; lowest=start
        for day,ids in days:
            ids=[i for i in ids if sfrac[i]>0 or len(ids)>1]
            b0=bank
            staked=0.0; ret=0.0
            # singles
            day_single=sum(sfrac[i] for i in ids)
            scale=min(1.0, CAP_DAY/day_single) if day_single>0 else 1.0
            outcomes={}
            for i in ids:
                p=(1/P[i][1]) if use_market else P[i][2]
                outcomes[i]=np.random.random()<p
            for i in ids:
                st=sfrac[i]*scale*b0
                staked+=st
                if outcomes[i]: ret+=st*P[i][1]
            # combo del día
            if len(ids)>=2:
                cst=COMBO_FRAC*b0; staked+=cst
                if all(outcomes[i] for i in ids):
                    co=1.0
                    for i in ids: co*=P[i][1]
                    ret+=cst*co
            net=ret-staked
            bank=b0+net
            if net>0:
                locked+=LOCK*net; bank-=LOCK*net
            lowest=min(lowest,bank+locked)
            if bank<=0: bank=0.0; break
        finals[t]=bank+locked
        minbank[t]=lowest
    return finals,minbank

for label,mk in [("SI TIENES RAZÓN (tus prob)",False),("SI EL MERCADO TIENE RAZÓN",True)]:
    f,mn=simulate(40000,mk)
    pct=lambda q: np.percentile(f,q)
    print(f"\n=== {label} ===  (banca inicial S/500)")
    print(f"  Mediana final:      S/ {np.median(f):,.0f}")
    print(f"  Media final:        S/ {f.mean():,.0f}")
    print(f"  P10 / P90:          S/ {pct(10):,.0f}  /  S/ {pct(90):,.0f}")
    print(f"  P(terminar en verde, >500):   {(f>500).mean()*100:,.1f}%")
    print(f"  P(perder >50% banca alguna vez): {(mn<250).mean()*100:,.1f}%")
    print(f"  P(banca casi a cero, <100):   {(f<100).mean()*100:,.1f}%")

print("\n=== STAKE % POR PICK (fracción de banca, ⅔ Kelly, cap 6%) ===")
for day,ids in days:
    s=", ".join(f"#{i} {sfrac[i]*100:.1f}%" for i in ids if sfrac[i]>0)
    print(f"{day}: {s}")
json.dump({i:round(sfrac[i],4) for i in P},open("sfrac.json","w"))
