# pod pythona 2

GroupA=['RUS']
GroupB=[]
GroupC=[]
GroupD=[]
GroupE=[]
GroupF=[]
GroupG=[]
GroupH=[]
Groups=(GroupA,GroupB,GroupC,GroupD,GroupE,GroupF,GroupG,GroupH)

UEFA=['RUS','BEL','GER','POL','POR','FRA','ENG','SPA','SUI','CRO','ICE','SWE','DEN','SER']
CONMEBOL=['BRA','ARG','URG','COL','PER']
CONCACAF=['MEX','CRC','PAN']
AFC=['IRN','JAP','KOR','KSA','AUS']
CAF=['NGA','TUN','MAR','SEN','EGY']

POT1=['BRA','BEL','GER','POL','POR','FRA','ARG']
POT2=['MEX','ENG','SPA','URG','COL','SUI','CRO','PER']
POT3=['ICE','SWE','DEN','IRN','CRC','EGY','TUN','SEN']
POT4=['NGA','KSA','AUS','JAP','MAR','KOR']
POTSER=['SER']
POTPAN=['PAN']


import random

#Koszyk 1
for y in Groups:
    if len(y)==0:
        pick=random.choice(POT1)
        y.append(pick)
        POT1.remove(pick)

#Koszyk 2 - jedynym ograniczeniem Ameryka Pd
for y in Groups:
    pick=random.choice(POT2)
    while pick in CONMEBOL and y[0] in CONMEBOL:
        pick=random.choice(POT2)
    y.append(pick)
    POT2.remove(pick)
    
#Koszyk3 - meksyk oraz max 2 europejskie

for y in Groups:
    if y[1]=='MEX':
        pick=random.choice(POT3)
        while pick=='CRC':
            pick=random.choice(POT3)
        y.append(pick)
        POT3.remove(pick)   
    if y[0] in UEFA and y[1] in UEFA:
        pick=random.choice(POT3)
        while pick in UEFA:
            pick=random.choice(POT3)
        y.append(pick)
        POT3.remove(pick)


for y in Groups:
    if len(y)==2:
        pick=random.choice(POT3)
        y.append(pick)
        POT3.remove(pick)

#Koszyk4 - iran, panama, serbia, afryka, dziabanina

#wywalamy serbow
GroupSER=[]

for y in Groups:
    if len(y)==3 and y[0] in UEFA:
        if y[1] not in UEFA:
            if y[2] not in UEFA:
                GroupSER.append(y)
    if len(y)==3 and y[0] not in UEFA:
        if y[1] in UEFA:
            if y[2] not in UEFA:
                GroupSER.append(y)
        else:
            GroupSER.append(y)
           
SERout=random.choice(GroupSER)
SERout.extend(POTSER)

#reszta koszyka
for y in Groups:

    if len(y)==3 and y[2] in CAF:
        pick=random.choice(POT4)
        while pick in CAF:
            pick=random.choice(POT4)
        y.append(pick)
        if pick in POT4:
            POT4.remove(pick)
            
    if len(y)==3 and y[2]=='IRN':
        pick=random.choice(POT4)
        while pick in AFC:
            pick=random.choice(POT4)
        y.append(pick)
        if pick in POT4:
            POT4.remove(pick)

#wywalamy panamczykow
GroupPAN=[]

for y in Groups:
    if len(y)==3 and y[1]!='MEX':
        if y[2]!='CRC':
            GroupPAN.append(y)

PANout=random.choice(GroupPAN)
PANout.extend(POTPAN)

#reszta
for y in Groups:
    if len(y)==3:
        pick=random.choice(POT4)
        y.append(pick)
        if pick in POT4:
            POT4.remove(pick)

print "A",GroupA
print "B",GroupB
print "C",GroupC
print "D",GroupD
print "E",GroupE
print "F",GroupF
print "G",GroupG
print "H",GroupH
