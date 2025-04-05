"""
Created on Sun Mar 30 11:10:14 2025

@author: borgar stenseth
e-post: borgar.stenseth@gmail.com

Dette er phyton filen som er laget for å løse den ferdig definerte prosjektoppgaven
"""
# %% Del A - les inn fra filen og generer 4 arrays i 4 kolonner


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Leser excel filen
sup_w_24 = pd.read_excel("support_uke_24.xlsx") # filnavnet er support_uke_24

#konverterer kolonner til NP arrays

u_dag = sup_w_24['Ukedag'].to_numpy() #Data fra excel files kolonne 1 legges inn i var uke dag
kl_slett = sup_w_24['Klokkeslett'].to_numpy()
varighet = sup_w_24['Varighet'].to_numpy()
score = sup_w_24['Tilfredshet'].to_numpy()


# %% Del B - Skrive ett program som finer antall hendvendelser gor hver ukeda i uken 24 og hvis det i ett plott

henv_man = np.sum(u_dag == 'Mandag')
henv_tir = np.sum(u_dag == 'Tirsdag')
henv_ons = np.sum(u_dag == 'Onsdag')
henv_tor = np.sum(u_dag == 'Torsdag')
henv_fre = np.sum(u_dag == 'Fredag')

print("/n === Utskrift av oppgave del b, dette blir også plottet i ett eget plott ===")
print("Antall hendvendelser på mandag er: ", henv_man, "tirsdag er:", henv_tir, "onsdag er:", henv_ons, "torsdag er:", henv_tor, "og fredag er:", henv_fre)

print("Se også plottet for grafisk fremstilling")

# Telle antall hendelser per dag

per_dag, ant_per_dag = np.unique(u_dag, return_counts=True)

# Definere rekkefølge på ukedagene
order = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag']

# Sortere resultatene etter ønsket rekkegølge

soterte_dager = sorted(zip(per_dag, ant_per_dag), key=lambda x: order.index(x[0]))

# Skrive ut resultatene 
for dag, antall in soterte_dager:
    print(f"{dag} var det {antall} support hendvendelser.")


# Sorter data i henhold til ønsket rekkefølge
sorterte_dager = [order.index(dag) for dag in per_dag] #Sortere slik at dagene blir man til fredag og ikke alfabetisk
sorteringsindeks = np.argsort(sorterte_dager)
per_dag_sortert = per_dag[sorteringsindeks]
ant_per_dag_sortert = ant_per_dag[sorteringsindeks]

# plotting av hendvendelser per dag
plt.close('all')
plt.bar(per_dag_sortert, ant_per_dag_sortert, color='skyblue')  # Bruk markører for klarhet
plt.xlabel('ukedager')
plt.ylabel('antall hendvendelser')
plt.grid(axis='y', color='pink', linestyle='--', linewidth=0.8, alpha=0.5)

plt.show()

# %% Del C - Skriv ett program som finner minste og lengste samtale

# Indeksene for lengste og korteste samtale
kl_slett_max = np.argmax(varighet)
kl_slett_max_var = kl_slett[kl_slett_max]
varighet_max = varighet[kl_slett_max]

kl_slett_min = np.argmin(varighet)
kl_slett_min_var = kl_slett[kl_slett_min]
varighet_min = varighet[kl_slett_min]

# Finne de ukedagene der det matches den lengste og korteste samtalen
u_dag_max = u_dag[kl_slett_max]
u_dag_min = u_dag[kl_slett_min]


print("\n=== Del C - Utskrift av den største og minste samtale varigheten ===")
print("Den lengste samtalen var klokken: ", kl_slett_max_var, "paa", u_dag_max, "og varte i: ", varighet_max, "hh:mm:ss")
print("Den korteste samtalen var klokken: ", kl_slett_min_var, "paa", u_dag_min,"og varte i: ", varighet_min, "hh:mm:ss")

# %% Del D - Skriv ett program som regner ut gjennomsnittelig samtaletid basert på alle hendvend i uke 24

# først ønsker jeg å gjøre om varighet kollonnen fra sting format til total tid i sekunder. 
# dette for enklere å regne ut middelverdi

# Konverter varighet til sekunder ved hjelp av listeforståelse
varighet_sek = np.array([
    sum(int(t) * 60**i for i, t in enumerate(reversed(tid.split(":"))))
    for tid in varighet])

mid_varighet_sek = np.mean(varighet_sek)

print(f"Middelverdi i sekunder: {mid_varighet_sek:.2f}") #printe ut i sekunder med to desimaler (:.2f )
print(f"Middelverdi i minutter: {mid_varighet_sek / 60:.2f}") # printe ut minutter med to desimaler

# Formatere tilbake til hh:mm:ss format
timer, resterende_sekunder = divmod(mid_varighet_sek, 3600)
minutter, sekunder = divmod(resterende_sekunder, 60)
tid_formatert = f"{int(timer):02}:{int(minutter):02}:{int(sekunder):02}"

print(f"Middelverdi i sekunder: {mid_varighet_sek:.4f}") #med fire desimaler
print(f"Middelverdi som hh:mm:ss: {tid_formatert}") #utskrift som viser orginalt tidsformat.

# %% Del e - Andtall hendvendelser per 2-timers bolk.

from datetime import datetime

kl_slett_konv = [datetime.strptime(element, "%H:%M:%S").time() for element in kl_slett]    #Konvertererr slik at vikan gjøre if statement og valg på klokkeslett

# Initialiserer tellere for de oppgitte tidsperiodene
ant_8_10 = 0
ant_10_12 = 0
ant_12_14 = 0 
ant_14_16 = 0

# Bruke en for løkke for å telle innenfor hvert tidsområde

for tids_omr in kl_slett_konv:
    if 8 <= tids_omr.hour <10:
        ant_8_10 += 1
    elif 10 <= tids_omr.hour <12:
        ant_10_12 += 1
    elif 12 <= tids_omr.hour <14:
        ant_12_14 +=1
    elif 14 <= tids_omr.hour <16:
        ant_14_16 +=1 
        
        
# Totalsummering
totalt_antall_1 = henv_man + henv_tir + henv_ons + henv_tor + henv_fre
totalt_antall_2 = ant_8_10 + ant_10_12 + ant_12_14 + ant_14_16
        
# Utskrift av antall per tidsområde

print("\n=== Oppgave del e - Utskrift av antall hendvendelser i intrevallene 08:00-10:00, 10.00 - 12.00, 12.00 - 14.00 og 14.00 - 16.00 ===")
print(f"Antallet hendvendelser mellom klokka 8 og 10 er: {ant_8_10} fordelt over 5 ukedager")
print(f"Antallet hendvendelser mellom klokka 10 og 12 er: {ant_10_12} fordelt over 5 ukedager")
print(f"Antallet hendvendelser mellom klokka 12 og 14 er: {ant_12_14} fordelt over 5 ukedager")
print(f"Antallet hendvendelser mellom klokka 14 og 16 er: {ant_14_16} fordelt over 5 ukedager")
print("\n=== Test og sammenligning at alle entries er tatt med. ===")
print(f"Som en test, totalt antall ved å summere hendvedelser per dag, som er: {totalt_antall_1} skal være lik summering av tidsperioder {totalt_antall_2}")

# %% Del f - kundens tilfredshet 

ant_neg = 0
ant_noyt = 0
ant_pos = 0

for ant_pos_noyt_neg in score:
    if 0 <= ant_pos_noyt_neg <=6:
        ant_neg += 1
    elif 7 <= ant_pos_noyt_neg <=8:
        ant_noyt += 1
    elif 9 <= ant_pos_noyt_neg <=10:
        ant_pos += 1
        
print(f"Antall negative hendvendelser er: {ant_neg}")        
print(f"Antall nøytrale hendvendelser er: {ant_noyt}")
print(f"Antall positive hendvendelser er: {ant_pos}")

# Totalsummering av antall som har gitt score
tot_ant_score = ant_pos + ant_noyt + ant_neg
print(f"Det totale antallet som har gitt en score er: {tot_ant_score}")

# Utregning av NPS

perc_prom = (ant_pos/tot_ant_score)
perc_pass = (ant_noyt/tot_ant_score)
perc_nega = (ant_neg/tot_ant_score)

NPS = round(((perc_prom - perc_nega)*100))

print("Antallet kunder som er positive er: ", round(perc_prom*100), "mens prosentent nøytrale er", round(perc_pass*100), "og prosent andel negative er: ", round(perc_nega*100))
print("Dette vil da gi en Net Promoter Score (NPS) på verdien: ", NPS)


