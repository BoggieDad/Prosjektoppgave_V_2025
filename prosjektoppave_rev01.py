# -*- coding: utf-8 -*-
"""
Created on Sun Mar 30 11:10:14 2025

@author: borgar stenseth
e-post: borgar.stenseth@gmail.com

Dette er phyton filen som er laget for å løse den ferdigdefinerte prosjektoppgaven
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

# Skriv ut NumPy-arrays
print("Kolonne 1:", u_dag)
print("Kolonne 2:", kl_slett)
print("Kolonne 3:", varighet)
print("Kolonne 4:", score)

# %% Del B - Skrive ett program som finer antall hendvendelser gor hver ukeda i uken 24 og hvis det i ett plott

henv_man = np.sum(u_dag == 'Mandag')
henv_tir = np.sum(u_dag == 'Tirsdag')
henv_ons = np.sum(u_dag == 'Onsdag')
henv_tor = np.sum(u_dag == 'Torsdag')
henv_fre = np.sum(u_dag == 'Fredag')
print("Utskrift av oppgave del b, dette blir også plottet i ett eget plott")
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

kl_slett_max = np.argmax(varighet)
varighet_max = varighet[kl_slett_max]
print("Den lengste samtalen var klokken: ", kl_slett_max, "og varte i: ", varighet_max)
print("klokkeslett", kl_slett)