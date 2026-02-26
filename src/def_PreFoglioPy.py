import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import io
import xlsxwriter
#from scipy.cluster.hierarchy import linkage, fcluster
from collections import defaultdict

def delta(listMax, listMin):
    delta = []
    for iMax, iMin in zip(listMax, listMin):
        delta_I = np.abs(np.abs(iMax) - np.abs(iMin))
        delta.append(delta_I)
    return delta

def importMidasData(path):
    #path: percorso dove trovare il file excel di input
    # if path == None:
    #     path = "\Out_Midas\00_Info_Modello.xlsx"

    # Leggere tutti i fogli in un dizionario
    # Le chiavi del dizionario sono i nomi dei fogli
    # Leggi solo i fogli specifici
    
    specific_sheets = ['Point', 'Element', 'CDS', 'Mobili']
    xls = path

    # Ora puoi accedere a ogni foglio come un DataFrame dal dizionario
    #for sheet_name, df in xls.items():
        #print(f"Sheet name: {sheet_name}")
        #print(df)

    # xls è ora un dizionario con i fogli 'point' ed 'element'
    # Accedere ai dati del foglio 'point'
    point_df = xls['Point']
    point_df = point_df.astype({"Node": int, "X": float, "Y": float, "Z": float})
    point_df.set_index('Node', inplace=True)
    #print("Dati del foglio 'point':")
    #print(point_df)

    # Accedere ai dati del foglio 'element'
    element_df = xls['Element']
    element_df = element_df.astype({"Element": float, "Material":float, "Property":float, "Node1":float, "Node2":float})
    element_df.set_index('Element', inplace=True)

    #print("\nDati del foglio 'element':")
    #print(element_df)

    # Dizionario per memorizzare i DataFrame filtrati
    filtered_dfs = {}
    filtered_dfs["Point"] = point_df.T.to_dict()
    filtered_dfs["Element"] = element_df.T.to_dict()

    # Accedere ai dati del foglio 'CDS'
    cds_df = xls['CDS']

    # Filtra il DataFrame per ogni valore unico in "Load"
    filtered_dfs["G1"] = cds_df[cds_df['Load'] == "G1"]
    filtered_dfs["G2"] = cds_df[cds_df['Load'] == "G2"]
    filtered_dfs["Ritiro"] = cds_df[cds_df['Load'] == "E2-Ritiro"]
    # Filtra per includere solo le righe dove "Load" è "'q7.1-Termica-'" o "'q7.1-Termica+'"
    filtered_dfs["Temperatura"] = cds_df[cds_df['Load'].isin(['Temperatura(max)', 'Temperatura(min)'])]
    filtered_dfs["Cedimenti"] = cds_df[cds_df['Load'].isin(['Cedimenti(max)', 'Cedimenti(min)'])]
    filtered_dfs["Varo"] = cds_df[cds_df['Load'].isin(['Varo(max)', 'Varo(min)'])]

    # Accedere ai dati del foglio 'Mobili'
    mobili_df = xls['Mobili']

    # Filtrare basandosi sulla prima lettera
    filtered_dfs["Tandem"] = mobili_df[mobili_df['Load'].apply(lambda x: x.startswith('T'))]
    filtered_dfs["Distr"] = mobili_df[mobili_df['Load'].apply(lambda x: x.startswith('D'))]
    filtered_dfs["Fatica"] = mobili_df[mobili_df['Load'].apply(lambda x: x.startswith('F'))]
    filtered_dfs["Vento"] = mobili_df[mobili_df['Load'].apply(lambda x: x.startswith('V'))]

    return filtered_dfs

def envelopeSLU(df_data):
    c_g1 = 1.35 #pesi propi
    c_g2 = 1.5 #permanenti portati
    c_mc = 1.35 #mobili concentrati
    c_md = 1.35 #mobili distribuiti
    c_r = 1.2 #ritiro
    c_c = 1.2 #cedimenti
    c_t = 1.5*0.6 #temperatura
    c_w = 1.5*0.6 #vento

    colonne_da_moltiplicare = ["Axial",  "Shear-y",  "Shear-z",  "Torsion",  "Moment-y",  "Moment-z"]
    g1 = df_data["G1"][colonne_da_moltiplicare].apply(lambda x: x * c_g1)
    g1["Elem"] = df_data["G1"]["Elem"]
    g1["Part"] = df_data["G1"]["Part"]

    g2 = df_data["G2"][colonne_da_moltiplicare].apply(lambda x: x * c_g2)
    g2["Elem"] = df_data["G2"]["Elem"]
    g2["Part"] = df_data["G2"]["Part"]

    r = df_data["Ritiro"][colonne_da_moltiplicare].apply(lambda x: x * c_r)
    r["Elem"] = df_data["Ritiro"]["Elem"]
    r["Part"] = df_data["Ritiro"]["Part"]

    c = df_data["Cedimenti"][colonne_da_moltiplicare].apply(lambda x: x * c_c)


    t = df_data["Temperatura"][colonne_da_moltiplicare].apply(lambda x: x * c_t)


    #w = df_data["Vento"][colonne_da_moltiplicare].apply(lambda x: x * c_w)
    mc = df_data["Tandem"][colonne_da_moltiplicare].apply(lambda x: x * c_mc)
    md = df_data["Distr"][colonne_da_moltiplicare].apply(lambda x: x * c_md)


    # Colonne da sommare
    columns_to_sum = ["Axial", "Shear-y", "Shear-z", "Torsion", "Moment-y", "Moment-z"]
    slu = pd.DataFrame()
    # Somma delle colonne specificate
    for col in columns_to_sum:
        slu[col] = g1[f'{col}_df1'] + g2[f'{col}_df2']

    slu = g1[colonne_da_moltiplicare] + g2[colonne_da_moltiplicare] #+ r[colonne_da_moltiplicare]


    print(g1)
    print(slu)



    return

def importOneLoad_MIDAS(df_Data):
    #path: percorso dove trovare il file excel di input

    dictLoad = df_Data.T.to_dict()

    dictLoad_order = {}
    for i in dictLoad:
        keys = list(dictLoad[i].keys())[2:]
        element = dictLoad[i]['Elem']

        try:
            dictLoad_order[element]
        except:
            dictLoad_order[element] = {}
            dictLoad_order[element]['I'] = {}
            dictLoad_order[element]['J'] = {}

        if dictLoad[i]['Part'][0] == 'I':
            for ikeys in keys:
                dictLoad_order[element]['I'][ikeys] = dictLoad[i][ikeys]
                if ikeys == 'Part':
                    pI = dictLoad_order[element]['I'][ikeys].replace('I[', "").replace(']', "")
                    dictLoad_order[element]['I'][ikeys] = int(pI)
                #print(element)

        elif dictLoad[i]['Part'][0] == 'J':
            for ikeys in keys:
                #print(ikeys)
                #print(dictLoad[i][ikeys])
                dictLoad_order[element]['J'][ikeys] = dictLoad[i][ikeys]
                if ikeys == 'Part':
                    pJ = dictLoad_order[element]['J'][ikeys].replace('J[', "").replace(']', "")
                    dictLoad_order[element]['J'][ikeys] = int(pJ)

    return dictLoad_order 

def importMultiLoad_MIDAS(df_Data):
    #path: percorso dove trovare il file excel di input

    unique_loads = df_Data['Load'].unique()

    dictMultiLoad = {}
    #for iFoglio in xl.sheet_names:  PRIMA
    for load in unique_loads:
        # Creazione di un nuovo DataFrame per ogni valore unico
        dictLoad = df_Data[df_Data['Load'] == load].T.to_dict()

        dictLoad_order = {'Axial': {}, 'Shear-z': {}, 'Moment-y': {}, 'Torsion': {}}
        for i in dictLoad:
            keys = list(dictLoad[i].keys())[2:]
            element = dictLoad[i]['Elem']
            refCDS = dictLoad[i]['Component']
            #print(element, refCDS)

            try:
                dictLoad_order[refCDS][element]
            except:
                dictLoad_order[refCDS][element] = {}
                dictLoad_order[refCDS][element]['I'] = {}
                dictLoad_order[refCDS][element]['J'] = {}

            if dictLoad[i]['Part'][0] == 'I':
                for ikeys in keys:
                    dictLoad_order[refCDS][element]['I'][ikeys] = dictLoad[i][ikeys]
                    if ikeys == 'Part':
                        pI = dictLoad_order[refCDS][element]['I'][ikeys].replace('I[', "").replace(']', "")
                        dictLoad_order[refCDS][element]['I'][ikeys] = int(pI)
                    #print(element)

            elif dictLoad[i]['Part'][0] == 'J':
                for ikeys in keys:
                    #print(ikeys)
                    #print(dictLoad[i][ikeys])
                    dictLoad_order[refCDS][element]['J'][ikeys] = dictLoad[i][ikeys]
                    if ikeys == 'Part':
                        pJ = dictLoad_order[refCDS][element]['J'][ikeys].replace('J[', "").replace(']', "")
                        dictLoad_order[refCDS][element]['J'][ikeys] = int(pJ)

        dictMultiLoad[load] = dictLoad_order

    return dictMultiLoad

def importMultiLoad2_MIDAS(df_Data):
    #path: percorso dove trovare il file excel di input
    unique_loads = df_Data['Load'].unique()

    # Element
    unique_loads = df_Data['Load'].unique()

    dictMultiLoad = {}
    for load in unique_loads:
        # Creazione di un nuovo DataFrame per ogni valore unico
        dictLoad = df_Data[df_Data['Load'] == load].T.to_dict()

        dictLoad_order = {}
        for i in dictLoad:
            keys = list(dictLoad[i].keys())[2:]
            element = dictLoad[i]['Elem']

            try:
                dictLoad_order[element]
            except:
                dictLoad_order[element] = {}
                dictLoad_order[element]['I'] = {}
                dictLoad_order[element]['J'] = {}

            if dictLoad[i]['Part'][0] == 'I':
                for ikeys in keys:
                    dictLoad_order[element]['I'][ikeys] = dictLoad[i][ikeys]
                    if ikeys == 'Part':
                        pI = dictLoad_order[element]['I'][ikeys].replace('I[', "").replace(']', "")
                        dictLoad_order[element]['I'][ikeys] = int(pI)
                    #print(element)

            elif dictLoad[i]['Part'][0] == 'J':
                for ikeys in keys:
                    #print(ikeys)
                    #print(dictLoad[i][ikeys])
                    dictLoad_order[element]['J'][ikeys] = dictLoad[i][ikeys]
                    if ikeys == 'Part':
                        pJ = dictLoad_order[element]['J'][ikeys].replace('J[', "").replace(']', "")
                        dictLoad_order[element]['J'][ikeys] = int(pJ)

        dictMultiLoad[load] = dictLoad_order

    return dictMultiLoad

def EleConcio(dictModel):

    dictConci = {}
    for i in dictModel['Element']:
        section = dictModel['Element'][i]['Property']
        if section >= 0:
            try:
                dictConci[section]['ele'].append(i)
            except:
                dictConci[section] = {'ele': [i]}

    #Per identificare i punti iniziali e finali dei conci
    for i in dictConci:
        ele = dictConci[i]['ele']
        #print("ele", dictConci[i]['ele'])
        coordI_X = []
        coordJ_X = []
        for j in ele:
            nodeI = dictModel['Element'][j]['Node1']
            nodeJ = dictModel['Element'][j]['Node2']
            coordI_X.append(dictModel['Point'][nodeI]['X'])
            coordJ_X.append(dictModel['Point'][nodeJ]['X'])

        maxI, maxJ = max(coordI_X), max(coordJ_X)
        minI, minJ = min(coordI_X), min(coordJ_X)
        if maxI >= maxJ:
            index = coordI_X.index(maxI)
            pointEnd = dictModel['Element'][ele[index]]['Node1']
            dictConci[i]['pointEnd'] = pointEnd
        elif maxI < maxJ:
            index = coordJ_X.index(maxJ)
            pointEnd = dictModel['Element'][ele[index]]['Node2']
            dictConci[i]['pointEnd'] = pointEnd

        if minI <= minJ:
            index = coordI_X.index(minI)
            pointStart = dictModel['Element'][ele[index]]['Node1']
            dictConci[i]['pointStart'] = pointStart
        elif minI > minJ:
            index = coordJ_X.index(minJ)
            pointStart = dictModel['Element'][ele[index]]['Node2']
            dictConci[i]['pointStart'] = pointStart
        
        dictConci[i]['Sollecitazioni'] = {'Forza Normale': {'G1+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'G1-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'G2+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'G2-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'R+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'R-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Mfat+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Mfat-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'MQ+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'MQ-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Md+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Md-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Mf+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Mf-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'T+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'T-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'C+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'C-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'V+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'V-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        },  

        'Momento flettente': {'G1+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'G1-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'G2+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'G2-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'R+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'R-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Mfat+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Mfat-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'MQ+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'MQ-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Md+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Md-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Mf+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Mf-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'T+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'T-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'C+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'C-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'V+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'V-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        }, 

        'Taglio': {'G1+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'G1-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'G2+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'G2-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'R+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'R-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Mfat+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Mfat-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'MQ+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'MQ-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Md+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Md-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Mf+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Mf-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'T+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'T-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'C+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'C-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'V+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'V-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        },  

        'Momento torcente': {'G1+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'G1-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'G2+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'G2-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'R+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'R-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Mfat+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Mfat-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'MQ+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'MQ-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Md+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Md-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'Mf+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'Mf-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'T+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'T-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 
        'C+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'C-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        'V+':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0}, 'V-':{'N_el': 0, 'N': 0.0, 'T': 0.0, 'Mf': 0.0, 'Mt': 0.0},
        } }

    return dictConci


def PlotConci(dictModel, dictConci):
    # Definizione del colormap
    colormap = plt.cm.get_cmap("Paired")
    num_keys = len(dictConci)

    # Creazione della mappa dei colori
    if num_keys > 1:
        color_map = {key: colormap(i / (num_keys - 1)) for i, key in enumerate(dictConci)}
    else:
        color_map = {key: colormap(0) for key in dictConci}
    
    # Iterazione su ogni concio
    for key, group in dictConci.items():
        line_color = color_map.get(key, 'black')  # Ottieni il colore per il concio corrente
        points = []  # Lista per memorizzare i punti di tutti gli elementi del concio
        
        # Iterazione sugli elementi del concio per raccogliere le coordinate dei nodi
        for element_id in group['ele']:
            node1, node2 = dictModel['Element'][element_id]['Node1'], dictModel['Element'][element_id]['Node2']
            points.append([dictModel['Point'][node1]['X'], dictModel['Point'][node1]['Y']])
            points.append([dictModel['Point'][node2]['X'], dictModel['Point'][node2]['Y']])
        
        # Conversione della lista dei punti in un array numpy per facilitare il calcolo
        points_array = np.array(points)
        
        # Calcolo del baricentro del concio corrente
        centroid_x = np.mean(points_array[:, 0])
        centroid_y = np.mean(points_array[:, 1])
        
        # Visualizzazione del nome della chiave al baricentro del concio
        #plt.text(centroid_x, centroid_y, key, color='black', fontsize=9, ha='center')
        
        # Disegno delle linee per gli elementi del concio
        for i in range(0, len(points), 2):
            plt.plot([points[i][0], points[i+1][0]], [points[i][1], points[i+1][1]], color=line_color)
    
    plt.axis('equal')
    plt.show()


def Plot_CDS(dictModel, dictLoad):
    
    for i in dictLoad:
        #print(dictLoad[i])
        #point I and J
        pI = dictLoad[i]['I']['Part']
        pJ = dictLoad[i]['J']['Part']
        #print(pI)
        #cds I
        Ni = dictLoad[i]['I']['Axial']
        Vyi = dictLoad[i]['I']['Shear-y']
        Vzi = dictLoad[i]['I']['Shear-z']
        Ti = dictLoad[i]['I']['Torsion']
        Myi = dictLoad[i]['I']['Moment-y']
        Mzi = dictLoad[i]['I']['Moment-z']
        #cds J
        Nj = dictLoad[i]['J']['Axial']
        Vyj = dictLoad[i]['J']['Shear-y']
        Vzj = dictLoad[i]['J']['Shear-z']
        Tj = dictLoad[i]['J']['Torsion']
        Myj = dictLoad[i]['J']['Moment-y']
        Mzj = dictLoad[i]['J']['Moment-z']
        #coordinate X point I and J
        Xi = dictModel['Point'][pI]['X']
        Xj = dictModel['Point'][pJ]['X']
        #Momento
        plt.plot( np.array([ Xi, Xj]), np.array([ Myi, Myj]), '-o')
    plt.show()
    return 

def Plot_CDS_concio(dictModel, dictLoad, dictConci):
    
    for j in dictConci:
        for i in dictConci[j]['ele']:
            #point I and J
            pI = dictLoad[i]['I']['Part']
            pJ = dictLoad[i]['J']['Part']
            #print(pI)
            #cds I
            Ni = dictLoad[i]['I']['Axial']
            Vyi = dictLoad[i]['I']['Shear-y']
            Vzi = dictLoad[i]['I']['Shear-z']
            Ti = dictLoad[i]['I']['Torsion']
            Myi = dictLoad[i]['I']['Moment-y']
            Mzi = dictLoad[i]['I']['Moment-z']
            #cds J
            Nj = dictLoad[i]['J']['Axial']
            Vyj = dictLoad[i]['J']['Shear-y']
            Vzj = dictLoad[i]['J']['Shear-z']
            Tj = dictLoad[i]['J']['Torsion']
            Myj = dictLoad[i]['J']['Moment-y']
            Mzj = dictLoad[i]['J']['Moment-z']
            #coordinate X point I and J
            Xi = dictModel['Point'][pI]['X']
            Xj = dictModel['Point'][pJ]['X']
        #Momento
            plt.plot( np.array([ Xi, Xj]), np.array([ Myi, Myj]), '-o')
    plt.show()
    return 

def AssignCDS_concio(dictModel, dictConci, dictLoad, NameCDS):
    cdsNameMax = NameCDS + '+'
    cdsNameMin = NameCDS + '-'
    for j in dictConci:
        N_I, V_I, M_I, T_I = [], [], [], []
        N_J, V_J, M_J, T_J = [], [], [], []
        for i in dictConci[j]['ele']:
            #point I and J
            #print(i)
            pI = dictLoad[i]['I']['Part']
            pJ = dictLoad[i]['J']['Part']
            #print(pI)
            #cds I
            Ni = dictLoad[i]['I']['Axial']
            N_I.append(Ni)
            Vyi = abs(dictLoad[i]['I']['Shear-y'])
            Vzi = abs(dictLoad[i]['I']['Shear-z'])
            V_I.append(Vzi)
            Ti = dictLoad[i]['I']['Torsion']
            T_I.append(Ti)
            Myi = dictLoad[i]['I']['Moment-y']
            M_I.append(Myi)
            Mzi = dictLoad[i]['I']['Moment-z']
            #cds J
            Nj = dictLoad[i]['J']['Axial']
            N_J.append(Nj)
            Vyj = abs(dictLoad[i]['J']['Shear-y'])
            Vzj = abs(dictLoad[i]['J']['Shear-z'])
            V_J.append(Vzj)
            Tj = dictLoad[i]['J']['Torsion']
            T_J.append(Tj)
            Myj = dictLoad[i]['J']['Moment-y']
            M_J.append(Myj)
            Mzj = dictLoad[i]['J']['Moment-z']
            #coordinate X point I and J
            Xi = dictModel['Point'][pI]['X']
            Xj = dictModel['Point'][pJ]['X']

            #Max and Min
            ####################################### Forza Normale
            Nmax = max(max(N_I), max(N_J))
            Nmin = min(min(N_I), min(N_J))
            try:
                indexMax = N_I.index(Nmax)
                T = V_I[indexMax]
                Mf = M_I[indexMax]
                Mt = T_I[indexMax]
            except:
                indexMax = N_J.index(Nmax)
                T = V_J[indexMax]
                Mf = M_J[indexMax]
                Mt = T_J[indexMax]
            
            N_ele = dictConci[j]['ele'][indexMax]
            dictConci[j]['Sollecitazioni']['Forza Normale'][cdsNameMax] = {'N_el': N_ele, 'N': Nmax, 'T': T, 'Mf': Mf, 'Mt': Mt}

            try:
                indexMax = N_I.index(Nmin)
                T = V_I[indexMax]
                Mf = M_I[indexMax]
                Mt = T_I[indexMax]
            except:
                indexMax = N_J.index(Nmin)
                T = V_J[indexMax]
                Mf = M_J[indexMax]
                Mt = T_J[indexMax]
            
            N_ele = dictConci[j]['ele'][indexMax]
            dictConci[j]['Sollecitazioni']['Forza Normale'][cdsNameMin] = {'N_el': N_ele, 'N': Nmin, 'T': T, 'Mf': Mf, 'Mt': Mt}


            ####################################### Taglio
            Vmax = max(max(V_I), max(V_J))
            #Vmin = min(min(V_I), min(V_J))

            try:
                indexMax = V_I.index(Vmax)

            except:
                indexMax = V_J.index(Vmax)

            
            Ni = N_I[indexMax]
            Mfi = M_I[indexMax]
            Mti = T_I[indexMax]

            Nj = N_J[indexMax]
            Mfj = M_J[indexMax]
            Mtj = T_J[indexMax]
            
            if Mfi > Mfj:
                N_ele = dictConci[j]['ele'][indexMax]
                dictConci[j]['Sollecitazioni']['Taglio'][cdsNameMax] = {'N_el': N_ele, 'N': Ni, 'T': Vmax, 'Mf': Mfi, 'Mt': Mtj}
                dictConci[j]['Sollecitazioni']['Taglio'][cdsNameMin] = {'N_el': N_ele, 'N': Nj, 'T': Vmax, 'Mf': Mfj, 'Mt': Mtj}
             
            elif Mfi < Mfj:
                N_ele = dictConci[j]['ele'][indexMax]
                dictConci[j]['Sollecitazioni']['Taglio'][cdsNameMin] = {'N_el': N_ele, 'N': Ni, 'T': Vmax, 'Mf': Mfi, 'Mt': Mtj}
                dictConci[j]['Sollecitazioni']['Taglio'][cdsNameMax] = {'N_el': N_ele, 'N': Nj, 'T': Vmax, 'Mf': Mfj, 'Mt': Mtj}
              
            
            ####################################### Momento flettente
            Mmax = max(max(M_I), max(M_J))
            Mmin = min(min(M_I), min(M_J))

            try:
                indexMax = M_I.index(Mmax)
                N = N_I[indexMax]
                T = V_I[indexMax]
                Mt = T_I[indexMax]
            except:
                indexMax = M_J.index(Mmax)
                N = N_J[indexMax]
                T = V_J[indexMax]
                Mt = T_J[indexMax]
            
            N_ele = dictConci[j]['ele'][indexMax]
            dictConci[j]['Sollecitazioni']['Momento flettente'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': T, 'Mf': Mmax, 'Mt': Mt}

            try:
                indexMax = M_I.index(Mmin)
                N = N_I[indexMax]
                T = V_I[indexMax]
                Mt = T_I[indexMax]
            except:
                indexMax = M_J.index(Mmin)
                N = N_J[indexMax]
                T = V_J[indexMax]
                Mt = T_J[indexMax]
            
            N_ele = dictConci[j]['ele'][indexMax]
            #print('cdsName', cdsNameMin)
            dictConci[j]['Sollecitazioni']['Momento flettente'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': T, 'Mf': Mmin, 'Mt': Mt}

            ####################################### Torsione
            Tmax = max(max(T_I), max(T_J))
            Tmin = min(min(T_I), min(T_J))

            try:
                indexMax = M_I.index(Mmax)
                N = N_I[indexMax]
                T = V_I[indexMax]
                Mf = M_I[indexMax]
            except:
                indexMax = M_J.index(Mmax)
                N = N_J[indexMax]
                T = V_J[indexMax]
                Mf = M_J[indexMax]
            
            N_ele = dictConci[j]['ele'][indexMax]
            dictConci[j]['Sollecitazioni']['Momento torcente'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': T, 'Mf': Mf, 'Mt': Tmax}

            try:
                indexMax = M_I.index(Mmin)
                N = N_I[indexMax]
                T = V_I[indexMax]
                Mf = M_I[indexMax]
            except:
                indexMax = M_J.index(Mmin)
                N = N_J[indexMax]
                T = V_J[indexMax]
                Mf = M_J[indexMax]
            
            N_ele = dictConci[j]['ele'][indexMax]
            dictConci[j]['Sollecitazioni']['Momento torcente'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': T, 'Mf': Mf, 'Mt': Tmin}

    return dictConci

def AssignCDSMulti_concio(dictModel, dictConci, dictLoad, NameCDS):
    ## Funzione con i file excel con più shet dentro
    cdsNameMax = NameCDS + '+'
    cdsNameMin = NameCDS + '-'

    nameFogli = list(dictLoad.keys())
    refCDS = list(dictLoad[nameFogli[0]].keys())

    
    dictMultiLoad = {}
    for i in dictConci:
        dictMultiLoad[i] = {'ele':dictConci[i]['ele'], 
        'Axial': { 'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}, 
        'Shear-z': { 'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}, 
        'Moment-y': { 'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}, 
        'Torsion': { 'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}, 
        }

    for iCDS in refCDS: # se non c'e la forza normale mettere refCDS[1:] 
        for iNameF in nameFogli:
            for j in dictConci:
                N_I, V_I, M_I, T_I = [], [], [], []
                N_J, V_J, M_J, T_J = [], [], [], []
                #print(dictConci[j]['ele'])
                for i in dictConci[j]['ele']:
                    #print(i)
                    
                    #point I and J
                    pI = dictLoad[iNameF][iCDS][i]['I']['Part']
                    pJ = dictLoad[iNameF][iCDS][i]['J']['Part']
                    #print(pI)
                    #cds I
                    Ni = dictLoad[iNameF][iCDS][i]['I']['Axial']
                    N_I.append(Ni)
                    Vyi = abs(dictLoad[iNameF][iCDS][i]['I']['Shear-y'])
                    Vzi = abs(dictLoad[iNameF][iCDS][i]['I']['Shear-z'])
                    V_I.append(Vzi)
                    Ti = dictLoad[iNameF][iCDS][i]['I']['Torsion']
                    T_I.append(Ti)
                    Myi = dictLoad[iNameF][iCDS][i]['I']['Moment-y']
                    M_I.append(Myi)
                    Mzi = dictLoad[iNameF][iCDS][i]['I']['Moment-z']
                    #cds J
                    Nj = dictLoad[iNameF][iCDS][i]['J']['Axial']
                    N_J.append(Nj)
                    Vyj = abs(dictLoad[iNameF][iCDS][i]['J']['Shear-y'])
                    Vzj = abs(dictLoad[iNameF][iCDS][i]['J']['Shear-z'])
                    V_J.append(Vzj)
                    Tj = dictLoad[iNameF][iCDS][i]['J']['Torsion']
                    T_J.append(Tj)
                    Myj = dictLoad[iNameF][iCDS][i]['J']['Moment-y']
                    M_J.append(Myj)
                    Mzj = dictLoad[iNameF][iCDS][i]['J']['Moment-z']
                    #coordinate X point I and J
                    Xi = dictModel['Point'][pI]['X']
                    Xj = dictModel['Point'][pJ]['X']

                dictMultiLoad[j][iCDS]['Axial']['I'].append(N_I)
                dictMultiLoad[j][iCDS]['Shear-z']['I'].append(V_I)
                dictMultiLoad[j][iCDS]['Moment-y']['I'].append(M_I)
                dictMultiLoad[j][iCDS]['Torsion']['I'].append(T_I)
                dictMultiLoad[j][iCDS]['Axial']['J'].append(N_J)
                dictMultiLoad[j][iCDS]['Shear-z']['J'].append(V_J)
                dictMultiLoad[j][iCDS]['Moment-y']['J'].append(M_J)
                dictMultiLoad[j][iCDS]['Torsion']['J'].append(T_J)
   
    for i in dictConci:
        #Max and Min
        ####################################### Forza Normale
        Nref_NI = np.array(dictMultiLoad[i]['Axial']['Axial']['I']).T
        Nref_NJ = np.array(dictMultiLoad[i]['Axial']['Axial']['J']).T

        VI = np.array(dictMultiLoad[i]['Axial']['Shear-z']['I']).T
        MI = np.array(dictMultiLoad[i]['Axial']['Moment-y']['I']).T
        TI = np.array(dictMultiLoad[i]['Axial']['Torsion']['I']).T

        VJ = np.array(dictMultiLoad[i]['Axial']['Shear-z']['J']).T
        MJ = np.array(dictMultiLoad[i]['Axial']['Moment-y']['J']).T
        TJ = np.array(dictMultiLoad[i]['Axial']['Torsion']['J']).T

        #INVILUPPO FORZA NORMALE
        NmaxInv_I = np.amax(Nref_NI).tolist()
        NmaxInv_J = np.amax(Nref_NJ).tolist()
        NminInv_I = np.amin(Nref_NI).tolist()
        NminInv_J = np.amin(Nref_NJ).tolist()


        Nmax = max(NmaxInv_I, NmaxInv_J)
        Nmin = min(NminInv_J, NminInv_J)
        # devo trovare due indici, il primo lo trovo dall'inviluppo e mi indica
        # quale elemento del concio è più sollecitato, 
        #poi devo trovare in che combinazione si trova

        try:
            res = np.where(Nref_NI == Nmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            V = VI[indexList][indexComb]
            Mf = MI[indexList][indexComb]
            Mt = TI[indexList][indexComb]
        except:
            res = np.where(Nref_NJ == Nmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            V = VJ[indexList][indexComb]
            Mf = MJ[indexList][indexComb]
            Mt = TJ[indexList][indexComb]
        
        #print('VJ', VJ)
        #print('ele', dictConci[j]['ele'])
        #print('index list', indexList)
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Forza Normale'][cdsNameMax] = {'N_el': N_ele, 'N': Nmin, 'T': V, 'Mf': Mf, 'Mt': Mt}

        try:
            res = np.where(Nref_NI == Nmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            V = VI[indexList][indexComb]
            Mf = MI[indexList][indexComb]
            Mt = TI[indexList][indexComb]
        except:
            res = np.where(Nref_NJ == Nmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            V = VJ[indexList][indexComb]
            Mf = MJ[indexList][indexComb]
            Mt = TJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Forza Normale'][cdsNameMin] = {'N_el': N_ele, 'N': Nmin, 'T': V, 'Mf': Mf, 'Mt': Mt}
        
        ####################################### Taglio
        Vref_VI = np.array(dictMultiLoad[i]['Shear-z']['Shear-z']['I']).T
        Vref_VJ = np.array(dictMultiLoad[i]['Shear-z']['Shear-z']['J']).T

        NI = np.array(dictMultiLoad[i]['Shear-z']['Axial']['I']).T
        MI = np.array(dictMultiLoad[i]['Shear-z']['Moment-y']['I']).T
        TI = np.array(dictMultiLoad[i]['Shear-z']['Torsion']['I']).T

        NJ = np.array(dictMultiLoad[i]['Shear-z']['Axial']['J']).T
        MJ = np.array(dictMultiLoad[i]['Shear-z']['Moment-y']['J']).T
        TJ = np.array(dictMultiLoad[i]['Shear-z']['Torsion']['J']).T

        #INVILUPPO TAGLIO
        VmaxInv_I = np.amax(Vref_VI).tolist()
        VmaxInv_J = np.amax(Vref_VJ).tolist()
        #VminInv_I = np.amin(Vref_VI).tolist()
        #VminInv_J = np.amin(Vref_VJ).tolist()


        Vmax = max(VmaxInv_I, VmaxInv_J)
        #Vmin = min(VminInv_J, VminInv_J)
        # devo trovare due indici, il primo lo trovo dall'inviluppo e mi indica
        # quale elemento del concio è più sollecitato, 
        #poi devo trovare in che combinazione si trova

        try:
            res = np.where(Vref_VI == Vmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            Mfi = [MI[indexList][indexcomb] for indexcomb in res[1]]
            Mfi_max = max(Mfi)
            Mfi_min = min(Mfi)
            indexCombi_Mfmax = MI[indexList].tolist().index(Mfi_max) # identifica in quale combinazione (lista trovare il max)
            indexCombi_Mfmin = MI[indexList].tolist().index(Mfi_min) # identifica in quale combinazione (lista trovare il min)  
        
            res = np.where(Vref_VJ == Vmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            Mfj = [MJ[indexList][indexcomb] for indexcomb in res[1]]
            Mfj_max = max(Mfj)
            Mfj_min = min(Mfj)
            indexCombj_Mfmax = MJ[indexList].tolist().index(Mfj_max) # identifica in quale combinazione (lista trovare il max)
            indexCombj_Mfmin = MJ[indexList].tolist().index(Mfj_min) # identifica in quale combinazione (lista trovare il min)  

            if Mfi_max >= Mfj_max:
                indexComb_max = indexCombi_Mfmax
                N_max = NI[indexList][indexComb_max]
                Mf_max = MI[indexList][indexComb_max]
                Mt_max = TI[indexList][indexComb_max]
                
            else:
                indexComb_max = indexCombj_Mfmax
                N_max = NJ[indexList][indexComb_max]
                Mf_max = MJ[indexList][indexComb_max]
                Mt_max = TJ[indexList][indexComb_max]
            
            if Mfi_min < Mfj_min:
                indexComb_min = indexCombi_Mfmin
                N_min = NI[indexList][indexComb_min]
                Mf_min = MI[indexList][indexComb_min]
                Mt_min = TI[indexList][indexComb_min]
            else:
                indexComb_min = indexCombj_Mfmin
                N_min = NJ[indexList][indexComb_min]
                Mf_min = MJ[indexList][indexComb_min]
                Mt_min = TJ[indexList][indexComb_min]

        except:
            #print('il taglio massimo non sta sulle stesse sezioni dell elemento')
            try:
                res = np.where(Vref_VI == Vmax)
                indexList = res[0][0] # per trovare l'elmento corrispondente
                Mfi = [MI[indexList][indexcomb] for indexcomb in res[1]]
                Mfi_max = max(Mfi)
                Mfi_min = min(Mfi)
                indexComb_max = MI[indexList].tolist().index(Mfi_max) # identifica in quale combinazione (lista trovare il max)
                indexComb_min = MI[indexList].tolist().index(Mfi_min) # identifica in quale combinazione (lista trovare il min)  
            
                N_max = NI[indexList][indexComb_max]
                N_min = NI[indexList][indexComb_min]
                Mf_max = MI[indexList][indexComb_max]
                Mf_min = MI[indexList][indexComb_min]
                Mt_max = TI[indexList][indexComb_max]
                Mt_min = TI[indexList][indexComb_min]
            except:
                res = np.where(Vref_VJ == Vmax)
                indexList = res[0][0] # per trovare l'elmento corrispondente
                Mfj = [MJ[indexList][indexcomb] for indexcomb in res[1]]
                Mfj_max = max(Mfj)
                Mfj_min = min(Mfj)
                indexComb_max = MJ[indexList].tolist().index(Mfj_max) # identifica in quale combinazione (lista trovare il max)
                indexComb_min = MJ[indexList].tolist().index(Mfj_min) # identifica in quale combinazione (lista trovare il min)  

                N_max = NJ[indexList][indexComb_max]
                N_min = NJ[indexList][indexComb_min]
                Mf_max = MJ[indexList][indexComb_max]
                Mf_min = MJ[indexList][indexComb_min]
                Mt_max = TJ[indexList][indexComb_max]
                Mt_min = TJ[indexList][indexComb_min]

        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Taglio'][cdsNameMax] = {'N_el': N_ele, 'N': N_max, 'T': Vmax, 'Mf': Mf_max, 'Mt': Mt_max}
        dictConci[i]['Sollecitazioni']['Taglio'][cdsNameMin] = {'N_el': N_ele, 'N': N_min, 'T': Vmax, 'Mf': Mf_min, 'Mt': Mt_min}
 
        ####################################### Momento flettente
        Mref_MI = np.array(dictMultiLoad[i]['Moment-y']['Moment-y']['I']).T
        Mref_MJ = np.array(dictMultiLoad[i]['Moment-y']['Moment-y']['J']).T

        NI = np.array(dictMultiLoad[i]['Moment-y']['Axial']['I']).T
        VI = np.array(dictMultiLoad[i]['Moment-y']['Shear-z']['I']).T
        TI = np.array(dictMultiLoad[i]['Moment-y']['Torsion']['I']).T

        NJ = np.array(dictMultiLoad[i]['Moment-y']['Axial']['J']).T
        VJ = np.array(dictMultiLoad[i]['Moment-y']['Shear-z']['J']).T
        TJ = np.array(dictMultiLoad[i]['Moment-y']['Torsion']['J']).T

        #INVILUPPO 
        MmaxInv_I = np.amax(Mref_MI).tolist()
        MmaxInv_J = np.amax(Mref_MJ).tolist()
        MminInv_I = np.amin(Mref_MI).tolist()
        MminInv_J = np.amin(Mref_MJ).tolist()


        Mmax = max(MmaxInv_I, MmaxInv_J)
        Mmin = min(MminInv_J, MminInv_J)
        # devo trovare due indici, il primo lo trovo dall'inviluppo e mi indica
        # quale elemento del concio è più sollecitato, 
        #poi devo trovare in che combinazione si trova

        try:
            res = np.where(Mref_MI == Mmax)
            #print('Ciao', res)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NI[indexList][indexComb]
            V = VI[indexList][indexComb]
            Mt = TI[indexList][indexComb]
        except:
            res = np.where(Mref_MJ == Mmax)
            #print('Ciao', res)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NJ[indexList][indexComb]
            V = VJ[indexList][indexComb]
            Mt = TJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        #print('elementi', dictConci[i]['ele'])
        #print('Momento max I', Mref_MI)
        #print('Momento max J', Mref_MJ)
        dictConci[i]['Sollecitazioni']['Momento flettente'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mmax, 'Mt': Mt}

        try:
            res = np.where(Mref_MI == Mmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NI[indexList][indexComb]
            V = VI[indexList][indexComb]
            Mt = TI[indexList][indexComb]
        except:
            res = np.where(Mref_MJ == Mmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NJ[indexList][indexComb]
            V = VJ[indexList][indexComb]
            Mt = TJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Momento flettente'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mmin, 'Mt': Mt}

        ####################################### Torsione
        Mtref_MtI = np.array(dictMultiLoad[i]['Torsion']['Torsion']['I']).T
        Mtref_MtJ = np.array(dictMultiLoad[i]['Torsion']['Torsion']['J']).T

        NI = np.array(dictMultiLoad[i]['Torsion']['Axial']['I']).T
        VI = np.array(dictMultiLoad[i]['Torsion']['Shear-z']['I']).T
        MI = np.array(dictMultiLoad[i]['Torsion']['Moment-y']['I']).T

        NJ = np.array(dictMultiLoad[i]['Torsion']['Axial']['J']).T
        VJ = np.array(dictMultiLoad[i]['Torsion']['Shear-z']['J']).T
        MJ = np.array(dictMultiLoad[i]['Torsion']['Moment-y']['J']).T

        #INVILUPPO 
        MtmaxInv_I = np.amax(Mtref_MtI).tolist()
        MtmaxInv_J = np.amax(Mtref_MtJ).tolist()
        MtminInv_I = np.amin(Mtref_MtI).tolist()
        MtminInv_J = np.amin(Mtref_MtJ).tolist()


        Tmax = max(MtmaxInv_I, MtmaxInv_J)
        Tmin = min(MtminInv_J, MtminInv_J)
        # devo trovare due indici, il primo lo trovo dall'inviluppo e mi indica
        # quale elemento del concio è più sollecitato, 
        #poi devo trovare in che combinazione si trova

        try:
            res = np.where(Mtref_MtI == Tmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NI[indexList][indexComb]
            V = VI[indexList][indexComb]
            Mf = MI[indexList][indexComb]
        except:
            res = np.where(Mtref_MtJ == Tmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NJ[indexList][indexComb]
            V = VJ[indexList][indexComb]
            Mf = MJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Momento torcente'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Tmax}

        try:
            res = np.where(Mtref_MtI == Tmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NI[indexList][indexComb]
            V = VI[indexList][indexComb]
            Mf = MI[indexList][indexComb]
        except:
            res = np.where(Mtref_MtJ == Tmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NJ[indexList][indexComb]
            V = VJ[indexList][indexComb]
            Mf = MJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Momento torcente'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Tmin}
    
    return dictConci

def AssignCDSMulti2_concio(dictModel, dictConci, dictLoad, NameCDS):
    ## Funzione con i file excel con più shet dentro
    cdsNameMax = NameCDS + '+'
    cdsNameMin = NameCDS + '-'

    nameFogli = list(dictLoad.keys())
    
    dictMultiLoad = {}
    for i in dictConci:
        dictMultiLoad[i] = {'ele':dictConci[i]['ele'], 
        'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}

    for iNameF in nameFogli:
        for j in dictConci:
            N_I, V_I, M_I, T_I = [], [], [], []
            N_J, V_J, M_J, T_J = [], [], [], []
            for i in dictConci[j]['ele']:
                #point I and J
                #pI = dictLoad[iNameF][i]['I']['Part']
                #pJ = dictLoad[iNameF][i]['J']['Part']
                #print(pI)
                #cds I
                Ni = dictLoad[iNameF][i]['I']['Axial']
                N_I.append(Ni)
                Vyi = abs(dictLoad[iNameF][i]['I']['Shear-y'])
                Vzi = abs(dictLoad[iNameF][i]['I']['Shear-z'])
                V_I.append(Vzi)
                Ti = dictLoad[iNameF][i]['I']['Torsion']
                T_I.append(Ti)
                Myi = dictLoad[iNameF][i]['I']['Moment-y']
                M_I.append(Myi)
                Mzi = dictLoad[iNameF][i]['I']['Moment-z']
                #cds J
                Nj = dictLoad[iNameF][i]['J']['Axial']
                N_J.append(Nj)
                Vyj = abs(dictLoad[iNameF][i]['J']['Shear-y'])
                Vzj = abs(dictLoad[iNameF][i]['J']['Shear-z'])
                V_J.append(Vzj)
                Tj = dictLoad[iNameF][i]['J']['Torsion']
                T_J.append(Tj)
                Myj = dictLoad[iNameF][i]['J']['Moment-y']
                M_J.append(Myj)
                Mzj = dictLoad[iNameF][i]['J']['Moment-z']
                #coordinate X point I and J
                #Xi = dictModel['Point'][pI]['X']
                #Xj = dictModel['Point'][pJ]['X']

            dictMultiLoad[j]['Axial']['I'].append(N_I)
            dictMultiLoad[j]['Shear-z']['I'].append(V_I)
            dictMultiLoad[j]['Moment-y']['I'].append(M_I)
            dictMultiLoad[j]['Torsion']['I'].append(T_I)
            dictMultiLoad[j]['Axial']['J'].append(N_J)
            dictMultiLoad[j]['Shear-z']['J'].append(V_J)
            dictMultiLoad[j]['Moment-y']['J'].append(M_J)
            dictMultiLoad[j]['Torsion']['J'].append(T_J)
   
    for i in dictConci:
        #Max and Min
        ####################################### Forza Normale
        Nref_NI = np.array(dictMultiLoad[i]['Axial']['I']).T
        Nref_NJ = np.array(dictMultiLoad[i]['Axial']['J']).T

        VI = np.array(dictMultiLoad[i]['Shear-z']['I']).T
        MI = np.array(dictMultiLoad[i]['Moment-y']['I']).T
        TI = np.array(dictMultiLoad[i]['Torsion']['I']).T

        VJ = np.array(dictMultiLoad[i]['Shear-z']['J']).T
        MJ = np.array(dictMultiLoad[i]['Moment-y']['J']).T
        TJ = np.array(dictMultiLoad[i]['Torsion']['J']).T

        #INVILUPPO
        NmaxInv_I = np.amax(Nref_NI).tolist()
        NmaxInv_J = np.amax(Nref_NJ).tolist()
        NminInv_I = np.amin(Nref_NI).tolist()
        NminInv_J = np.amin(Nref_NJ).tolist()


        Nmax = max(NmaxInv_I, NmaxInv_J)
        Nmin = min(NminInv_J, NminInv_J)
        # devo trovare due indici, il primo lo trovo dall'inviluppo e mi indica
        # quale elemento del concio è più sollecitato, 
        #poi devo trovare in che combinazione si trova

        try:
            res = np.where(Nref_NI == Nmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            V = VI[indexList][indexComb]
            Mf = MI[indexList][indexComb]
            Mt = TI[indexList][indexComb]
        except:
            res = np.where(Nref_NJ == Nmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            V = VJ[indexList][indexComb]
            Mf = MJ[indexList][indexComb]
            Mt = TJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Forza Normale'][cdsNameMax] = {'N_el': N_ele, 'N': Nmin, 'T': V, 'Mf': Mf, 'Mt': Mt}

        try:
            res = np.where(Nref_NI == Nmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            V = VI[indexList][indexComb]
            Mf = MI[indexList][indexComb]
            Mt = TI[indexList][indexComb]
        except:
            res = np.where(Nref_NJ == Nmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            V = VJ[indexList][indexComb]
            Mf = MJ[indexList][indexComb]
            Mt = TJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Forza Normale'][cdsNameMin] = {'N_el': N_ele, 'N': Nmin, 'T': V, 'Mf': Mf, 'Mt': Mt}
        
        
        ####################################### Taglio
        Vref_VI = np.array(dictMultiLoad[i]['Shear-z']['I']).T
        Vref_VJ = np.array(dictMultiLoad[i]['Shear-z']['J']).T

        NI = np.array(dictMultiLoad[i]['Axial']['I']).T
        MI = np.array(dictMultiLoad[i]['Moment-y']['I']).T
        TI = np.array(dictMultiLoad[i]['Torsion']['I']).T

        NJ = np.array(dictMultiLoad[i]['Axial']['J']).T
        MJ = np.array(dictMultiLoad[i]['Moment-y']['J']).T
        TJ = np.array(dictMultiLoad[i]['Torsion']['J']).T

        #INVILUPPO TAGLIO
        VmaxInv_I = np.amax(Vref_VI).tolist()
        VmaxInv_J = np.amax(Vref_VJ).tolist()
        #VminInv_I = np.amin(Vref_VI).tolist()
        #VminInv_J = np.amin(Vref_VJ).tolist()


        Vmax = max(VmaxInv_I, VmaxInv_J)
        #Vmin = min(VminInv_J, VminInv_J)
        # devo trovare due indici, il primo lo trovo dall'inviluppo e mi indica
        # quale elemento del concio è più sollecitato, 
        #poi devo trovare in che combinazione si trova

        try:
            res = np.where(Vref_VI == Vmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            Mfi = [MI[indexList][indexcomb] for indexcomb in res[1]]
            Mfi_max = max(Mfi)
            Mfi_min = min(Mfi)
            indexCombi_Mfmax = MI[indexList].tolist().index(Mfi_max) # identifica in quale combinazione (lista trovare il max)
            indexCombi_Mfmin = MI[indexList].tolist().index(Mfi_min) # identifica in quale combinazione (lista trovare il min)  
        
            res = np.where(Vref_VJ == Vmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            Mfj = [MJ[indexList][indexcomb] for indexcomb in res[1]]
            Mfj_max = max(Mfj)
            Mfj_min = min(Mfj)
            indexCombj_Mfmax = MJ[indexList].tolist().index(Mfj_max) # identifica in quale combinazione (lista trovare il max)
            indexCombj_Mfmin = MJ[indexList].tolist().index(Mfj_min) # identifica in quale combinazione (lista trovare il min)  

            if Mfi_max >= Mfj_max:
                indexComb_max = indexCombi_Mfmax
                N_max = NI[indexList][indexComb_max]
                Mf_max = MI[indexList][indexComb_max]
                Mt_max = TI[indexList][indexComb_max]
                
            else:
                indexComb_max = indexCombj_Mfmax
                N_max = NJ[indexList][indexComb_max]
                Mf_max = MJ[indexList][indexComb_max]
                Mt_max = TJ[indexList][indexComb_max]
            
            if Mfi_min < Mfj_min:
                indexComb_min = indexCombi_Mfmin
                N_min = NI[indexList][indexComb_min]
                Mf_min = MI[indexList][indexComb_min]
                Mt_min = TI[indexList][indexComb_min]
            else:
                indexComb_min = indexCombj_Mfmin
                N_min = NJ[indexList][indexComb_min]
                Mf_min = MJ[indexList][indexComb_min]
                Mt_min = TJ[indexList][indexComb_min]

        except:
            print('il taglio massimo non sta sulle stesse sezioni dell elemento')
            try:
                res = np.where(Vref_VI == Vmax)
                indexList = res[0][0] # per trovare l'elmento corrispondente
                Mfi = [MI[indexList][indexcomb] for indexcomb in res[1]]
                Mfi_max = max(Mfi)
                Mfi_min = min(Mfi)
                indexComb_max = MI[indexList].tolist().index(Mfi_max) # identifica in quale combinazione (lista trovare il max)
                indexComb_min = MI[indexList].tolist().index(Mfi_min) # identifica in quale combinazione (lista trovare il min)  
            
                N_max = NI[indexList][indexComb_max]
                N_min = NI[indexList][indexComb_min]
                Mf_max = MI[indexList][indexComb_max]
                Mf_min = MI[indexList][indexComb_min]
                Mt_max = TI[indexList][indexComb_max]
                Mt_min = TI[indexList][indexComb_min]
            except:
                res = np.where(Vref_VJ == Vmax)
                indexList = res[0][0] # per trovare l'elmento corrispondente
                Mfj = [MJ[indexList][indexcomb] for indexcomb in res[1]]
                Mfj_max = max(Mfj)
                Mfj_min = min(Mfj)
                indexComb_max = MJ[indexList].tolist().index(Mfj_max) # identifica in quale combinazione (lista trovare il max)
                indexComb_min = MJ[indexList].tolist().index(Mfj_min) # identifica in quale combinazione (lista trovare il min)  

                N_max = NJ[indexList][indexComb_max]
                N_min = NJ[indexList][indexComb_min]
                Mf_max = MJ[indexList][indexComb_max]
                Mf_min = MJ[indexList][indexComb_min]
                Mt_max = TJ[indexList][indexComb_max]
                Mt_min = TJ[indexList][indexComb_min]

        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Taglio'][cdsNameMax] = {'N_el': N_ele, 'N': N_max, 'T': Vmax, 'Mf': Mf_max, 'Mt': Mt_max}
        dictConci[i]['Sollecitazioni']['Taglio'][cdsNameMin] = {'N_el': N_ele, 'N': N_min, 'T': Vmax, 'Mf': Mf_min, 'Mt': Mt_min}

        ####################################### Momento flettente
        Mref_MI = np.array(dictMultiLoad[i]['Moment-y']['I']).T
        Mref_MJ = np.array(dictMultiLoad[i]['Moment-y']['J']).T

        NI = np.array(dictMultiLoad[i]['Axial']['I']).T
        VI = np.array(dictMultiLoad[i]['Shear-z']['I']).T
        TI = np.array(dictMultiLoad[i]['Torsion']['I']).T

        NJ = np.array(dictMultiLoad[i]['Axial']['J']).T
        VJ = np.array(dictMultiLoad[i]['Shear-z']['J']).T
        TJ = np.array(dictMultiLoad[i]['Torsion']['J']).T

        #INVILUPPO 
        MmaxInv_I = np.amax(Mref_MI).tolist()
        MmaxInv_J = np.amax(Mref_MJ).tolist()
        MminInv_I = np.amin(Mref_MI).tolist()
        MminInv_J = np.amin(Mref_MJ).tolist()


        Mmax = max(MmaxInv_I, MmaxInv_J)
        Mmin = min(MminInv_J, MminInv_J)
        # devo trovare due indici, il primo lo trovo dall'inviluppo e mi indica
        # quale elemento del concio è più sollecitato, 
        #poi devo trovare in che combinazione si trova

        try:
            res = np.where(Mref_MI == Mmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NI[indexList][indexComb]
            V = VI[indexList][indexComb]
            Mt = TI[indexList][indexComb]
        except:
            res = np.where(Mref_MJ == Mmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NJ[indexList][indexComb]
            V = VJ[indexList][indexComb]
            Mt = TJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Momento flettente'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mmax, 'Mt': Mt}

        try:
            res = np.where(Mref_MI == Mmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NI[indexList][indexComb]
            V = VI[indexList][indexComb]
            Mt = TI[indexList][indexComb]
        except:
            res = np.where(Mref_MJ == Mmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NJ[indexList][indexComb]
            V = VJ[indexList][indexComb]
            Mt = TJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Momento flettente'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mmin, 'Mt': Mt}

        ####################################### Torsione
        Mtref_MtI = np.array(dictMultiLoad[i]['Torsion']['I']).T
        Mtref_MtJ = np.array(dictMultiLoad[i]['Torsion']['J']).T

        NI = np.array(dictMultiLoad[i]['Axial']['I']).T
        VI = np.array(dictMultiLoad[i]['Shear-z']['I']).T
        MI = np.array(dictMultiLoad[i]['Moment-y']['I']).T

        NJ = np.array(dictMultiLoad[i]['Axial']['J']).T
        VJ = np.array(dictMultiLoad[i]['Shear-z']['J']).T
        MJ = np.array(dictMultiLoad[i]['Moment-y']['J']).T

        #INVILUPPO 
        MtmaxInv_I = np.amax(Mtref_MtI).tolist()
        MtmaxInv_J = np.amax(Mtref_MtJ).tolist()
        MtminInv_I = np.amin(Mtref_MtI).tolist()
        MtminInv_J = np.amin(Mtref_MtJ).tolist()


        Tmax = max(MtmaxInv_I, MtmaxInv_J)
        Tmin = min(MtminInv_J, MtminInv_J)
        # devo trovare due indici, il primo lo trovo dall'inviluppo e mi indica
        # quale elemento del concio è più sollecitato, 
        #poi devo trovare in che combinazione si trova

        try:
            res = np.where(Mtref_MtI == Tmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NI[indexList][indexComb]
            V = VI[indexList][indexComb]
            Mf = MI[indexList][indexComb]
        except:
            res = np.where(Mtref_MtJ == Tmax)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NJ[indexList][indexComb]
            V = VJ[indexList][indexComb]
            Mf = MJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Momento torcente'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Tmax}

        try:
            res = np.where(Mtref_MtI == Tmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NI[indexList][indexComb]
            V = VI[indexList][indexComb]
            Mf = MI[indexList][indexComb]
        except:
            res = np.where(Mtref_MtJ == Tmin)
            indexList = res[0][0] # per trovare l'elmento corrispondente
            indexComb = res[1][0] # identifica in quale combinazione (lista trovare il max) 
            N = NJ[indexList][indexComb]
            V = VJ[indexList][indexComb]
            Mf = MJ[indexList][indexComb]
        
        N_ele = dictConci[i]['ele'][indexList]
        dictConci[i]['Sollecitazioni']['Momento torcente'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Tmin}
    
    return dictConci

def AssignCDSFatica_concio(dictModel, dictConci, dictLoad, NameCDS):
    ## Funzione con i file excel con più shet dentro
    cdsNameMax = NameCDS + '+'
    cdsNameMin = NameCDS + '-'

    #print(dictLoad)
    nameFogli = list(dictLoad.keys())
    refCDS = list(dictLoad[nameFogli[0]].keys())

    
    dictMultiLoad = {}
    for i in dictConci:
        dictMultiLoad[i] = {'ele':dictConci[i]['ele'], 
        'Axial': { 'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}, 
        'Shear-z': { 'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}, 
        'Moment-y': { 'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}, 
        'Torsion': { 'Axial': {'I': [], 'J': []}, 'Shear-z': {'I': [], 'J': []}, 'Moment-y': {'I': [], 'J': []},  'Torsion': {'I': [], 'J': []}}, 
        }

    for iCDS in refCDS:
        for iNameF in nameFogli:
            for j in dictConci:
                N_I, V_I, M_I, T_I = [], [], [], []
                N_J, V_J, M_J, T_J = [], [], [], []
                for i in dictConci[j]['ele']:
                    #point I and J
                    pI = dictLoad[iNameF][iCDS][i]['I']['Part']
                    pJ = dictLoad[iNameF][iCDS][i]['J']['Part']
                    #print(pI)
                    #cds I
                    Ni = dictLoad[iNameF][iCDS][i]['I']['Axial']
                    N_I.append(Ni)
                    Vyi = dictLoad[iNameF][iCDS][i]['I']['Shear-y']
                    Vzi = dictLoad[iNameF][iCDS][i]['I']['Shear-z']
                    V_I.append(Vzi)
                    #print('vzi', Ni)
                    Ti = dictLoad[iNameF][iCDS][i]['I']['Torsion']
                    T_I.append(Ti)
                    Myi = dictLoad[iNameF][iCDS][i]['I']['Moment-y']
                    M_I.append(Myi)
                    Mzi = dictLoad[iNameF][iCDS][i]['I']['Moment-z']
                    #cds J
                    Nj = dictLoad[iNameF][iCDS][i]['J']['Axial']
                    N_J.append(Nj)
                    Vyj = dictLoad[iNameF][iCDS][i]['J']['Shear-y']
                    Vzj = dictLoad[iNameF][iCDS][i]['J']['Shear-z']
                    V_J.append(Vzj)
                    Tj = dictLoad[iNameF][iCDS][i]['J']['Torsion']
                    T_J.append(Tj)
                    Myj = dictLoad[iNameF][iCDS][i]['J']['Moment-y']
                    M_J.append(Myj)
                    Mzj = dictLoad[iNameF][iCDS][i]['J']['Moment-z']
                    #coordinate X point I and J
                    Xi = dictModel['Point'][pI]['X']
                    Xj = dictModel['Point'][pJ]['X']

                dictMultiLoad[j][iCDS]['Axial']['I'].append(N_I)
                dictMultiLoad[j][iCDS]['Shear-z']['I'].append(V_I)
                dictMultiLoad[j][iCDS]['Moment-y']['I'].append(M_I)
                dictMultiLoad[j][iCDS]['Torsion']['I'].append(T_I)
                dictMultiLoad[j][iCDS]['Axial']['J'].append(N_J)
                dictMultiLoad[j][iCDS]['Shear-z']['J'].append(V_J)
                dictMultiLoad[j][iCDS]['Moment-y']['J'].append(M_J)
                dictMultiLoad[j][iCDS]['Torsion']['J'].append(T_J)
   
    for i in dictConci:
        #Max and Min
        ####################################### Forza Normale
        NI = np.array(dictMultiLoad[i]['Axial']['Axial']['I']).T
        NJ = np.array(dictMultiLoad[i]['Axial']['Axial']['J']).T

        VI = np.array(dictMultiLoad[i]['Axial']['Shear-z']['I']).T
        MI = np.array(dictMultiLoad[i]['Axial']['Moment-y']['I']).T
        TI = np.array(dictMultiLoad[i]['Axial']['Torsion']['I']).T

        VJ = np.array(dictMultiLoad[i]['Axial']['Shear-z']['J']).T
        MJ = np.array(dictMultiLoad[i]['Axial']['Moment-y']['J']).T
        TJ = np.array(dictMultiLoad[i]['Axial']['Torsion']['J']).T

        #INVILUPPO
        NmaxInv_I = np.amax(NI, axis=1).tolist()  #'min value of every Row: '
        NmaxInv_J = np.amax(NJ, axis=1).tolist()
        NminInv_I = np.amin(NI, axis=1).tolist()
        NminInv_J = np.amin(NJ, axis=1).tolist()

        MmaxInv_I = np.amax(MI, axis=1).tolist()
        MmaxInv_J = np.amax(MJ, axis=1).tolist()
        MminInv_I = np.amin(MI, axis=1).tolist()
        MminInv_J = np.amin(MJ, axis=1).tolist()

        TmaxInv_I = np.amax(TI, axis=1).tolist()
        TmaxInv_J = np.amax(TJ, axis=1).tolist()
        TminInv_I = np.amin(TI, axis=1).tolist()
        TminInv_J = np.amin(TJ, axis=1).tolist()

        VmaxInv_I = np.amax(VI, axis=1).tolist()
        VmaxInv_J = np.amax(VJ, axis=1).tolist()
        VminInv_I = np.amin(VI, axis=1).tolist()
        VminInv_J = np.amin(VJ, axis=1).tolist()

        deltaI = delta(NmaxInv_I, NminInv_I)
        deltaJ = delta(NmaxInv_J, NminInv_J)

        deltaMax = max(max(deltaI), max(deltaJ))
        try:
            res = np.where(deltaI == deltaMax)
            index = deltaI.index(deltaMax)
            V = VmaxInv_I[index]
            N = NmaxInv_I[index]
            Mf = MmaxInv_I[index]
            Mt = TmaxInv_I[index]
        except:
            index = deltaJ.index(deltaMax)
            V = VmaxInv_J[index]
            N = NmaxInv_J[index]
            Mf = MmaxInv_J[index]
            Mt = TmaxInv_J[index]           
        
        N_ele = dictConci[i]['ele'][index]
        dictConci[i]['Sollecitazioni']['Forza Normale'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt}

        try:
            index = deltaI.index(deltaMax)
            V = VminInv_I[index]
            N = NminInv_I[index]
            Mf = MminInv_I[index]
            Mt = TminInv_I[index]
        except:
            index = deltaJ.index(deltaMax)
            V = VminInv_J[index]
            N = NminInv_J[index]
            Mf = MminInv_J[index]
            Mt = TminInv_J[index]           
        
        N_ele = dictConci[i]['ele'][index]
        dictConci[i]['Sollecitazioni']['Forza Normale'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt}
        
        
        ####################################### Taglio
        VI = np.array(dictMultiLoad[i]['Shear-z']['Shear-z']['I']).T
        VJ = np.array(dictMultiLoad[i]['Shear-z']['Shear-z']['J']).T

        NI = np.array(dictMultiLoad[i]['Shear-z']['Axial']['I']).T
        MI = np.array(dictMultiLoad[i]['Shear-z']['Moment-y']['I']).T
        TI = np.array(dictMultiLoad[i]['Shear-z']['Torsion']['I']).T

        NJ = np.array(dictMultiLoad[i]['Shear-z']['Axial']['J']).T
        MJ = np.array(dictMultiLoad[i]['Shear-z']['Moment-y']['J']).T
        TJ = np.array(dictMultiLoad[i]['Shear-z']['Torsion']['J']).T

        #INVILUPPO 
        VmaxInv_I = np.amax(VI, axis=1).tolist()
        VmaxInv_J = np.amax(VJ, axis=1).tolist()
        VminInv_I = np.amin(VI, axis=1).tolist()
        VminInv_J = np.amin(VJ, axis=1).tolist()

        MmaxInv_I = np.amax(MI, axis=1).tolist()
        MmaxInv_J = np.amax(MJ, axis=1).tolist()
        MminInv_I = np.amin(MI, axis=1).tolist()
        MminInv_J = np.amin(MJ, axis=1).tolist()

        TmaxInv_I = np.amax(TI, axis=1).tolist()
        TmaxInv_J = np.amax(TJ, axis=1).tolist()
        TminInv_I = np.amin(TI, axis=1).tolist()
        TminInv_J = np.amin(TJ, axis=1).tolist()

        NmaxInv_I = np.amax(NI, axis=1).tolist()
        NmaxInv_J = np.amax(NJ, axis=1).tolist()
        NminInv_I = np.amin(NI, axis=1).tolist()
        NminInv_J = np.amin(NJ, axis=1).tolist()

        deltaI = delta(VmaxInv_I, VminInv_I)
        deltaJ = delta(VmaxInv_J, VminInv_J)

        deltaMax = max(max(deltaI), max(deltaJ))
        try:
            
            #res = np.where(deltaI == deltaMax)
            #print(res)
            index = deltaI.index(deltaMax)
            V = VmaxInv_I[index]
            N = NmaxInv_I[index]
            Mf = MmaxInv_I[index]
            Mt = TmaxInv_I[index]
        except:
            index = deltaJ.index(deltaMax)
            V = VmaxInv_J[index]
            N = NmaxInv_J[index]
            Mf = MmaxInv_J[index]
            Mt = TmaxInv_J[index]           
        
        N_ele = dictConci[i]['ele'][index]
        dictConci[i]['Sollecitazioni']['Taglio'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt}

        try:
            index = deltaI.index(deltaMax)
            V = VminInv_I[index]
            N = NminInv_I[index]
            Mf = MminInv_I[index]
            Mt = TminInv_I[index]
        except:
            index = deltaJ.index(deltaMax)
            V = VminInv_J[index]
            N = NminInv_J[index]
            Mf = MminInv_J[index]
            Mt = TminInv_J[index]           
        
        N_ele = dictConci[i]['ele'][index]
        dictConci[i]['Sollecitazioni']['Taglio'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt}

        ####################################### Momento flettente
        MI = np.array(dictMultiLoad[i]['Moment-y']['Moment-y']['I']).T
        MJ = np.array(dictMultiLoad[i]['Moment-y']['Moment-y']['J']).T

        NI = np.array(dictMultiLoad[i]['Moment-y']['Axial']['I']).T
        VI = np.array(dictMultiLoad[i]['Moment-y']['Shear-z']['I']).T
        TI = np.array(dictMultiLoad[i]['Moment-y']['Torsion']['I']).T

        NJ = np.array(dictMultiLoad[i]['Moment-y']['Axial']['J']).T
        VJ = np.array(dictMultiLoad[i]['Moment-y']['Shear-z']['J']).T
        TJ = np.array(dictMultiLoad[i]['Moment-y']['Torsion']['J']).T

        #INVILUPPO 
        VmaxInv_I = np.amax(VI, axis=1).tolist()
        VmaxInv_J = np.amax(VJ, axis=1).tolist()
        VminInv_I = np.amin(VI, axis=1).tolist()
        VminInv_J = np.amin(VJ, axis=1).tolist()

        MmaxInv_I = np.amax(MI, axis=1).tolist()
        MmaxInv_J = np.amax(MJ, axis=1).tolist()
        MminInv_I = np.amin(MI, axis=1).tolist()
        MminInv_J = np.amin(MJ, axis=1).tolist()

        

        TmaxInv_I = np.amax(TI, axis=1).tolist()
        TmaxInv_J = np.amax(TJ, axis=1).tolist()
        TminInv_I = np.amin(TI, axis=1).tolist()
        TminInv_J = np.amin(TJ, axis=1).tolist()

        NmaxInv_I = np.amax(NI, axis=1).tolist()
        NmaxInv_J = np.amax(NJ, axis=1).tolist()
        NminInv_I = np.amin(NI, axis=1).tolist()
        NminInv_J = np.amin(NJ, axis=1).tolist()

        deltaI = delta(MmaxInv_I, MminInv_I)
        deltaJ = delta(MmaxInv_J, MminInv_J)
        #print(deltaJ)

        deltaMax = max(max(deltaI), max(deltaJ))
        try:
            index = deltaI.index(deltaMax)
            V = VmaxInv_I[index]
            N = NmaxInv_I[index]
            Mf = MmaxInv_I[index]
            #print(Mf)
            Mt = TmaxInv_I[index]
        except:
            index = deltaJ.index(deltaMax)
            V = VmaxInv_J[index]
            N = NmaxInv_J[index]
            Mf = MmaxInv_J[index]
            #print(Mf)
            Mt = TmaxInv_J[index]           
        
        N_ele = dictConci[i]['ele'][index]
        #print(i, cdsNameMax, {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt})
        dictConci[i]['Sollecitazioni']['Momento flettente'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt}

        try:
            index = deltaI.index(deltaMax)
            V = VminInv_I[index]
            N = NminInv_I[index]
            Mf = MminInv_I[index]
            Mt = TminInv_I[index]
        except:
            index = deltaJ.index(deltaMax)
            V = VminInv_J[index]
            N = NminInv_J[index]
            Mf = MminInv_J[index]
            Mt = TminInv_J[index]           
        
        N_ele = dictConci[i]['ele'][index]
        dictConci[i]['Sollecitazioni']['Momento flettente'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt}

        ####################################### Torsione
        TI = np.array(dictMultiLoad[i]['Torsion']['Torsion']['I']).T
        TJ = np.array(dictMultiLoad[i]['Torsion']['Torsion']['J']).T

        NI = np.array(dictMultiLoad[i]['Torsion']['Axial']['I']).T
        VI = np.array(dictMultiLoad[i]['Torsion']['Shear-z']['I']).T
        MI = np.array(dictMultiLoad[i]['Torsion']['Moment-y']['I']).T

        NJ = np.array(dictMultiLoad[i]['Torsion']['Axial']['J']).T
        VJ = np.array(dictMultiLoad[i]['Torsion']['Shear-z']['J']).T
        MJ = np.array(dictMultiLoad[i]['Torsion']['Moment-y']['J']).T

        #INVILUPPO 
        VmaxInv_I = np.amax(VI, axis=1).tolist()
        VmaxInv_J = np.amax(VJ, axis=1).tolist()
        VminInv_I = np.amin(VI, axis=1).tolist()
        VminInv_J = np.amin(VJ, axis=1).tolist()

        MmaxInv_I = np.amax(MI, axis=1).tolist()
        MmaxInv_J = np.amax(MJ, axis=1).tolist()
        MminInv_I = np.amin(MI, axis=1).tolist()
        MminInv_J = np.amin(MJ, axis=1).tolist()

        TmaxInv_I = np.amax(TI, axis=1).tolist()
        TmaxInv_J = np.amax(TJ, axis=1).tolist()
        TminInv_I = np.amin(TI, axis=1).tolist()
        TminInv_J = np.amin(TJ, axis=1).tolist()

        NmaxInv_I = np.amax(NI, axis=1).tolist()
        NmaxInv_J = np.amax(NJ, axis=1).tolist()
        NminInv_I = np.amin(NI, axis=1).tolist()
        NminInv_J = np.amin(NJ, axis=1).tolist()

        deltaI = delta(TmaxInv_I, TminInv_I)
        deltaJ = delta(TmaxInv_J, TminInv_J)

        deltaMax = max(max(deltaI), max(deltaJ))
        
        try:
            index = deltaI.index(deltaMax)
            V = VmaxInv_I[index]
            N = NmaxInv_I[index]
            Mf = MmaxInv_I[index]
            Mt = TmaxInv_I[index]
        except:
            index = deltaJ.index(deltaMax)
            V = VmaxInv_J[index]
            N = NmaxInv_J[index]
            Mf = MmaxInv_J[index]
            Mt = TmaxInv_J[index]           
        
        N_ele = dictConci[i]['ele'][index]
        dictConci[i]['Sollecitazioni']['Momento torcente'][cdsNameMax] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt}

        try:
            index = deltaI.index(deltaMax)
            V = VminInv_I[index]
            N = NminInv_I[index]
            Mf = MminInv_I[index]
            Mt = TminInv_I[index]
        except:
            index = deltaJ.index(deltaMax)
            V = VminInv_J[index]
            N = NminInv_J[index]
            Mf = MminInv_J[index]
            Mt = TminInv_J[index]           
        
        N_ele = dictConci[i]['ele'][index]
        dictConci[i]['Sollecitazioni']['Momento torcente'][cdsNameMin] = {'N_el': N_ele, 'N': N, 'T': V, 'Mf': Mf, 'Mt': Mt}
    
    return dictConci

def writeOut_xlsx(dictConci, NameFile=None):
    # Cretae a xlsx file
# --- GESTIONE OUTPUT (Disco o Memoria) ---
    if NameFile is None:
        # Modalità Streamlit (In Memoria)
        output_buffer = io.BytesIO()
        xlsx_File = xlsxwriter.Workbook(output_buffer, {'in_memory': True})
    else:
        # Modalità Locale (Su Disco)
        output_buffer = None
        xlsx_File = xlsxwriter.Workbook(NameFile)

    # Add new worksheet
    sheet_days = xlsx_File.add_worksheet()

    row1 = 3
    column = 1

    cell_format1 = xlsx_File.add_format({'bold': True, 'font_color': '#000000', 'bg_color': '#00FF00'})
    cell_format2 = xlsx_File.add_format({'bold': True, 'font_color': '#000000', 'bg_color': '#CCFFCC'})
    cell_format3 = xlsx_File.add_format({'bold': False, 'font_color': '#000000', 'bg_color': '#FFFF99'})

    sorted_keys = sorted(list(dictConci.keys()))# Ordina in ordine crescente
    #print("ciao", sorted_keys)
    
    for i in sorted_keys:
        sheet_days.write(row1, 0, 'Sezione numero:', cell_format1) #write numero sezione
        sheet_days.write(row1, 1, i, cell_format1) #write numero sezione
        for c in range(2, 6):
            sheet_days.write(row1, c, " ", cell_format1) #write numero sezione
        row1 += 1
        sheet_days.write(row1, 0, 'File') #write 
        sheet_days.write(row1, 1, 'N_el') #write 
        sheet_days.write(row1, 2, 'N') #write 
        sheet_days.write(row1, 3, 'T') #write 
        sheet_days.write(row1, 4, 'Mf') #write 
        sheet_days.write(row1, 5, 'Mt') #write 
        row1 += 1

        for iName in list(dictConci[i]['Sollecitazioni'].keys()):
            sheet_days.write(row1, 0, iName, cell_format2)
            for c in range(1, 6):
                sheet_days.write(row1, c, " ", cell_format2) #write numero sezione
            row1 += 1

            for iCDS in dictConci[i]['Sollecitazioni'][iName]:
                sheet_days.write(row1, 0, iCDS, cell_format3)

                for c in range(1, 6):
                    sheet_days.write(row1, c, " ", cell_format3) #write numero sezione

                row1 += 1
                try:
                    sheet_days.write(row1, 0, " ", cell_format3) #write 
                    sheet_days.write(row1, 1, dictConci[i]['Sollecitazioni'][iName][iCDS]['N_el'], cell_format3) #write N_ele
                    sheet_days.write(row1, 2, dictConci[i]['Sollecitazioni'][iName][iCDS]['N'], cell_format3) #write N
                    sheet_days.write(row1, 3, dictConci[i]['Sollecitazioni'][iName][iCDS]['T'], cell_format3) #write T
                    sheet_days.write(row1, 4, dictConci[i]['Sollecitazioni'][iName][iCDS]['Mf'], cell_format3) #write Mf
                    sheet_days.write(row1, 5, dictConci[i]['Sollecitazioni'][iName][iCDS]['Mt'], cell_format3) #write Mt
                except:
                    print('non sono stati assegnati le CDS in', iName, 'combinazione', iCDS, 'concio', i)
                row1 += 1
        
        row1 += 2

    # --- CHIUSURA E RITORNO ---
    xlsx_File.close()
    
    if output_buffer:
        # Se siamo in Streamlit, torniamo i bytes
        return output_buffer.getvalue()
    else:
        # Se siamo in locale, non torniamo nulla (il file è già su disco)
        return None

# ✅ function returns new dictionary (does NOT mutate original)

def remove_nested_keys(dictionary, keys_to_remove):
    new_dict = {}

    for key, value in dictionary.items():
        if key not in keys_to_remove:
            if isinstance(value, dict):
                new_dict[key] = remove_nested_keys(value, keys_to_remove)
            else:
                new_dict[key] = value

    return new_dict

def Run_Export1Out_SuperFoglio(input_data):
    #path: della cartella che contiene i file excel
    """
    input_dfs: è un dizionario contenente i fogli dell'excel caricato.
               Es: {'CDS': df_cds, 'Mobili': df_mobili, ...}
    """
    
# 1. ACQUISIZIONE DATI
    # Se input_data è un dizionario (da Streamlit)
    if isinstance(input_data, dict):
        df_cds = input_data.get('CDS')
        df_element = input_data.get('Element')
    else:
        # Fallback se passi un path (vecchio metodo)
        df_cds = pd.read_excel(input_data, sheet_name='CDS')
        # ... altri caricamenti
    
    if df_cds is None:
            return None
    #IMPORTAZIONE 
    print("oK - 0")
    dictModel = importMidasData(input_data)
    print("oK - 1")
    dictConci = EleConcio(dictModel)
    print("oK - 2")

    #### G1-Permanenti
    try:
        dictLoad_g1 = importOneLoad_MIDAS(dictModel["G1"])
        print("oK - 3")
        dictConci = AssignCDS_concio(dictModel, dictConci, dictLoad_g1, 'G1')
        print("oK - 4")
    except:
        print("G1 No Exists")

    #### G2-Permanenti
    try:
        dictLoad_g2 = importOneLoad_MIDAS(dictModel["G2"])
        dictConci = AssignCDS_concio(dictModel, dictConci, dictLoad_g2, 'G2')
    except:
        print("G2 No Exists")

    #### R-Ritiro
    try:
        dictLoad_R = importOneLoad_MIDAS(dictModel["Ritiro"])
        dictConci = AssignCDS_concio(dictModel, dictConci, dictLoad_R, 'R')
    except:
        print("Ritiro No Exists")

    #### V-VentoS
    #try:
    dictLoad_V = importMultiLoad_MIDAS(dictModel["Vento"])
    dictConci = AssignCDSMulti_concio(dictModel, dictConci, dictLoad_V, 'Mf')
    #except:
        #print("Vento No Exists")

    #### MQ - Mobili Tandem 
    #try:
    dictLoad_ts = importMultiLoad_MIDAS(dictModel["Tandem"])
    dictConci = AssignCDSMulti_concio(dictModel, dictConci, dictLoad_ts, 'MQ')
    #except:
        #print("Tandem No Exists")


    #### MQ - Mobili distribuiti
    try:
        dictLoad_udl = importMultiLoad_MIDAS(dictModel["Distr"])
        dictConci = AssignCDSMulti_concio(dictModel, dictConci, dictLoad_udl, 'Md')
    except:
        print("Distribuiti No Exists")

    #### T - Temperatura
    try:
        dictLoad_temp = importMultiLoad2_MIDAS(dictModel["Temperatura"])
        dictConci = AssignCDSMulti2_concio(dictModel, dictConci, dictLoad_temp, 'T')
    except:
        print("Temperatura.xlsx No Exists")

    #### C - Cedimenti
    try:
        dictLoad_c = importMultiLoad2_MIDAS(dictModel["Cedimenti"])
        dictConci = AssignCDSMulti2_concio(dictModel, dictConci, dictLoad_c, 'C')
    except:
        print("File 07_Cedimenti.xlsx No Exists")

    #### Mf - Fatica
    #if ("06_Fatica.xlsx" in fileList):
        #print("File 06_Fatica.xlsx Exists")
        #dictLoad_fatica = importMultiLoad_MIDAS(os.path.join(pathInput, "06_Fatica.xlsx"))
        #dictConci = AssignCDSFatica_concio(dictModel, dictConci, dictLoad_fatica, 'Mfat')
    #else:
        #print("File 06_fatica.xlsx No Exists")
    #devo lavorare sul massimo delta e non sul massimo della sollecitazione

    NewDict = remove_nested_keys(dictConci, ['Mfat+', 'Mfat-', 'V+', 'V-'])
    #NewDict = remove_nested_keys(dictConci, ['V+', 'V-'])


    res = writeOut_xlsx(NewDict) #, os.path.join(pathOut, "Output_GaudiCoseNTC_2023.xlsx")) 

    return res 

def Run_Export2Out_SuperFoglio(input_dfs, metodo = 2):
    #path: della cartella che contiene i file excel

# 1. Recupero dati
    if 'Mobili' in input_dfs:
        df_mobili = input_dfs['Mobili'].copy()
    else:
        return None
    
    #IMPORTAZIONE 
    dictModel = importMidasData(input_dfs)
    dictConci = EleConcio(dictModel)

    #### G1-Permanenti
    #try:
    #    dictLoad_g1 = importOneLoad_MIDAS(dictModel["G1"])
    #    dictConci = AssignCDS_concio(dictModel, dictConci, dictLoad_g1, 'G1')
    #except:
    #    print("G1 No Exists")

    #PlotConci(dictModel, dictConci)
    #print(dictModel['Point'])
    #print(dictModel['Element'])

    #print(dictConci[1]['Sollecitazioni'])
    #Plot_CDS(dictModel, dictLoad_g1)
    #Plot_CDS_concio(dictModel, dictLoad_g1, dictConci)

    if metodo == 1: #prende le cds in base alla variazione massima di sollecitazione
    #### Mf - Fatica
        try:
            dictLoad_fatica = importMultiLoad_MIDAS(dictModel["Fatica"])
            dictConci = AssignCDSFatica_concio(dictModel, dictConci, dictLoad_fatica, 'Mfat')
        except:
            print("fatica No Exists")
        #devo lavorare sul massimo delta e non sul massimo della sollecitazione

    elif metodo == 2: #prende CDS massima e minima come per i carihi mobili
    #### MQ - Mobili distribuiti
        try:
            dictLoad_udl = importMultiLoad_MIDAS(dictModel["Fatica"])
            dictConci = AssignCDSMulti_concio(dictModel, dictConci, dictLoad_udl, 'Mfat')
        except:
            print("Fatica No Exists")

    NewDict = remove_nested_keys(dictConci, ['Mf+', 'Mf-', 'V+', 'V-'])
    
    res = writeOut_xlsx(NewDict)

    return  res

def Run_Export3Out_SuperFoglio(pathInput): ##VARO
    #path: della cartella che contiene i file excel
    
    #IMPORTAZIONE 
    dictModel = importMidasData(pathInput)
    dictConci = EleConcio(dictModel)

    #### G1-Permanenti
    try:
        dictLoad_varo = importMultiLoad2_MIDAS(dictModel["Varo"])
        dictConci = AssignCDSMulti2_concio(dictModel, dictConci, dictLoad_varo, 'G1')
    except:
        print("G1 No Exists")

    #### Mf - Fatica
    #if ("06_Fatica.xlsx" in fileList):
        #print("File 06_Fatica.xlsx Exists")
        #dictLoad_fatica = importMultiLoad_MIDAS(os.path.join(pathInput, "06_Fatica.xlsx"))
        #dictConci = AssignCDSFatica_concio(dictModel, dictConci, dictLoad_fatica, 'Mfat')
    #else:
        #print("File 06_fatica.xlsx No Exists")
    #devo lavorare sul massimo delta e non sul massimo della sollecitazione

    NewDict = remove_nested_keys(dictConci, ['Mfat+', 'Mfat-', 'V+', 'V-'])
    #NewDict = remove_nested_keys(dictConci, ['V+', 'V-'])


    res = writeOut_xlsx(NewDict) #os.path.join(pathOut, "Output_GaudiCoseVaro_2023.xlsx")

    return res

def RunPlot(pathInput):
    #path: della cartella che contiene i file excel
    
    #IMPORTAZIONE 
    dictModel = importMidasData(pathInput)
    dictConci = EleConcio(dictModel)
    PlotConci(dictModel, dictConci)

    return



