import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import xlsxwriter

## import le funzioni di def_PreFoglioPy
import sys
from def_PreFoglioPy import * 


def inviluppoCDS_Static(dictLoad): #utilizzato ad esempio per la temperatura
    dictLoad_Inv = {}
    for iFoglio in list(dictLoad.keys()):
        for iEle in list(dictLoad[iFoglio].keys()):
            PartI = dictLoad[iFoglio][iEle]['I']['Part']
            PartJ = dictLoad[iFoglio][iEle]['J']['Part']

            NI = dictLoad[iFoglio][iEle]['I']['Axial']
            VyI = dictLoad[iFoglio][iEle]['I']['Shear-y']
            VzI = dictLoad[iFoglio][iEle]['I']['Shear-z']
            TI = dictLoad[iFoglio][iEle]['I']['Torsion']
            MyI = dictLoad[iFoglio][iEle]['I']['Moment-y']
            MzI = dictLoad[iFoglio][iEle]['I']['Moment-z']

            NJ = dictLoad[iFoglio][iEle]['J']['Axial']
            VyJ = dictLoad[iFoglio][iEle]['J']['Shear-y']
            VzJ = dictLoad[iFoglio][iEle]['J']['Shear-z']
            TJ = dictLoad[iFoglio][iEle]['J']['Torsion']
            MyJ = dictLoad[iFoglio][iEle]['J']['Moment-y']
            MzJ = dictLoad[iFoglio][iEle]['J']['Moment-z']

            try:
                dictLoad_Inv[iEle]['I']['Axial'].append(NI)
                dictLoad_Inv[iEle]['I']['Shear-y'].append(VyI)
                dictLoad_Inv[iEle]['I']['Shear-z'].append(VzI)
                dictLoad_Inv[iEle]['I']['Torsion'].append(TI)
                dictLoad_Inv[iEle]['I']['Moment-y'].append(MyI)
                dictLoad_Inv[iEle]['I']['Moment-z'].append(MzI)

                dictLoad_Inv[iEle]['J']['Axial'].append(NJ)
                dictLoad_Inv[iEle]['J']['Shear-y'].append(VyJ)
                dictLoad_Inv[iEle]['J']['Shear-z'].append(VzJ)
                dictLoad_Inv[iEle]['J']['Torsion'].append(TJ)
                dictLoad_Inv[iEle]['J']['Moment-y'].append(MyJ)
                dictLoad_Inv[iEle]['J']['Moment-z'].append(MzJ)

            except:
                dictLoad_Inv[iEle] = {'I': {'Part': PartI, 'Axial': [NI], 'Shear-y': [VyI], 'Shear-z': [VzI], 'Torsion': [TI], 'Moment-y': [MyI], 'Moment-z': [MzI]}, 
                                      'J': {'Part': PartJ, 'Axial': [NJ], 'Shear-y': [VyJ], 'Shear-z': [VzJ], 'Torsion': [TJ], 'Moment-y': [MyJ], 'Moment-z': [MzJ]}}

    
    # calculate Max and Min
    #calculate Max and Min
    dictMaxMin = {}
    #print(dictLoad_Inv.keys())
    Ref_cds = ['Axial', 'Shear-z', 'Torsion', 'Moment-y'] 
    for iCDS in Ref_cds:
        dictMaxMin[iCDS] = {'max': {}, 'min': {}}
        #print(iCDS)
        for iEle in dictLoad_Inv: #calcolo il max ed il min per ogni sezione

            valorMAX_I = max(dictLoad_Inv[iEle]['I'][iCDS])
            valorMAX_J = max(dictLoad_Inv[iEle]['J'][iCDS])
            valorMIN_I = min(dictLoad_Inv[iEle]['I'][iCDS])
            valorMIN_J = min(dictLoad_Inv[iEle]['J'][iCDS])

            indexMaxI = dictLoad_Inv[iEle]['I'][iCDS].index(valorMAX_I)
            indexMaxJ = dictLoad_Inv[iEle]['J'][iCDS].index(valorMAX_J)
            indexMinI = dictLoad_Inv[iEle]['I'][iCDS].index(valorMIN_I)
            indexMinJ = dictLoad_Inv[iEle]['J'][iCDS].index(valorMIN_J)

            #max
            NI = dictLoad_Inv[iEle]['I']['Axial'][indexMaxI]
            VyI = dictLoad_Inv[iEle]['I']['Shear-y'][indexMaxI]
            VzI = dictLoad_Inv[iEle]['I']['Shear-z'][indexMaxI]
            TI = dictLoad_Inv[iEle]['I']['Torsion'][indexMaxI]
            MyI = dictLoad_Inv[iEle]['I']['Moment-y'][indexMaxI]
            MzI = dictLoad_Inv[iEle]['I']['Moment-z'][indexMaxI]

            NJ = dictLoad_Inv[iEle]['J']['Axial'][indexMaxJ]
            VyJ = dictLoad_Inv[iEle]['J']['Shear-y'][indexMaxJ]
            VzJ = dictLoad_Inv[iEle]['J']['Shear-z'][indexMaxJ]
            TJ = dictLoad_Inv[iEle]['J']['Torsion'][indexMaxJ]
            MyJ = dictLoad_Inv[iEle]['J']['Moment-y'][indexMaxJ]
            MzJ = dictLoad_Inv[iEle]['J']['Moment-z'][indexMaxJ]

            dictMaxMin[iCDS]['max'][iEle] =  {'I': {'Part': PartI, 'Axial': NI, 'Shear-y': VyI, 'Shear-z': VzI, 'Torsion': TI, 'Moment-y': MyI, 'Moment-z': MzI}, 
                                            'J': {'Part': PartJ, 'Axial': NJ, 'Shear-y': VyJ, 'Shear-z': VzJ, 'Torsion': TJ, 'Moment-y': MyJ, 'Moment-z': MzJ}}


            #min
            NI = dictLoad_Inv[iEle]['I']['Axial'][indexMinI]
            VyI = dictLoad_Inv[iEle]['I']['Shear-y'][indexMinI]
            VzI = dictLoad_Inv[iEle]['I']['Shear-z'][indexMinI]
            TI = dictLoad_Inv[iEle]['I']['Torsion'][indexMinI]
            MyI = dictLoad_Inv[iEle]['I']['Moment-y'][indexMinI]
            MzI = dictLoad_Inv[iEle]['I']['Moment-z'][indexMinI]

            NJ = dictLoad_Inv[iEle]['J']['Axial'][indexMinJ]
            VyJ = dictLoad_Inv[iEle]['J']['Shear-y'][indexMinJ]
            VzJ = dictLoad_Inv[iEle]['J']['Shear-z'][indexMinJ]
            TJ = dictLoad_Inv[iEle]['J']['Torsion'][indexMinJ]
            MyJ = dictLoad_Inv[iEle]['J']['Moment-y'][indexMinJ]
            MzJ = dictLoad_Inv[iEle]['J']['Moment-z'][indexMinJ]

            dictMaxMin[iCDS]['min'][iEle] = {'I': {'Part': PartI, 'Axial': NI, 'Shear-y': VyI, 'Shear-z': VzI, 'Torsion': TI, 'Moment-y': MyI, 'Moment-z': MzI}, 
                                            'J': {'Part': PartJ, 'Axial': NJ, 'Shear-y': VyJ, 'Shear-z': VzJ, 'Torsion': TJ, 'Moment-y': MyJ, 'Moment-z': MzJ}}


    return dictLoad_Inv, dictMaxMin #, dictMaxMin

def inviluppoCDS_MoveLoad(dictLoad): #utilizzato per i mobili

    dictLoad_Inv = {}
    for iCDS in list(dictLoad[list(dictLoad.keys())[0]].keys()):
        dictLoad_Inv[iCDS] = { }

    for iFoglio in list(dictLoad.keys()):
        for iCDS in list(dictLoad[iFoglio].keys()):
            for iEle in list(dictLoad[iFoglio][iCDS].keys()):

                PartI = dictLoad[iFoglio][iCDS][iEle]['I']['Part']
                PartJ = dictLoad[iFoglio][iCDS][iEle]['J']['Part']

                NI = dictLoad[iFoglio][iCDS][iEle]['I']['Axial']
                VyI = dictLoad[iFoglio][iCDS][iEle]['I']['Shear-y']
                VzI = dictLoad[iFoglio][iCDS][iEle]['I']['Shear-z']
                TI = dictLoad[iFoglio][iCDS][iEle]['I']['Torsion']
                MyI = dictLoad[iFoglio][iCDS][iEle]['I']['Moment-y']
                MzI = dictLoad[iFoglio][iCDS][iEle]['I']['Moment-z']

                NJ = dictLoad[iFoglio][iCDS][iEle]['J']['Axial']
                VyJ = dictLoad[iFoglio][iCDS][iEle]['J']['Shear-y']
                VzJ = dictLoad[iFoglio][iCDS][iEle]['J']['Shear-z']
                TJ = dictLoad[iFoglio][iCDS][iEle]['J']['Torsion']
                MyJ = dictLoad[iFoglio][iCDS][iEle]['J']['Moment-y']
                MzJ = dictLoad[iFoglio][iCDS][iEle]['J']['Moment-z']

                try:
                    dictLoad_Inv[iCDS][iEle]['I']['Axial'].append(NI)
                    dictLoad_Inv[iCDS][iEle]['I']['Shear-y'].append(VyI)
                    dictLoad_Inv[iCDS][iEle]['I']['Shear-z'].append(VzI)
                    dictLoad_Inv[iCDS][iEle]['I']['Torsion'].append(TI)
                    dictLoad_Inv[iCDS][iEle]['I']['Moment-y'].append(MyI)
                    dictLoad_Inv[iCDS][iEle]['I']['Moment-z'].append(MzI)

                    dictLoad_Inv[iCDS][iEle]['J']['Axial'].append(NJ)
                    dictLoad_Inv[iCDS][iEle]['J']['Shear-y'].append(VyJ)
                    dictLoad_Inv[iCDS][iEle]['J']['Shear-z'].append(VzJ)
                    dictLoad_Inv[iCDS][iEle]['J']['Torsion'].append(TJ)
                    dictLoad_Inv[iCDS][iEle]['J']['Moment-y'].append(MyJ)
                    dictLoad_Inv[iCDS][iEle]['J']['Moment-z'].append(MzJ)

                except:
                    
                    dictLoad_Inv[iCDS][iEle] = {'I': {'Part': PartI, 'Axial': [NI], 'Shear-y': [VyI], 'Shear-z': [VzI], 'Torsion': [TI], 'Moment-y': [MyI], 'Moment-z': [MzI]}, 
                                                'J': {'Part': PartJ, 'Axial': [NJ], 'Shear-y': [VyJ], 'Shear-z': [VzJ], 'Torsion': [TJ], 'Moment-y': [MyJ], 'Moment-z': [MzJ]}}


    #calculate Max and Min
    dictMaxMin = {}
    #print(dictLoad_Inv.keys())
    for iCDS in dictLoad_Inv:
        dictMaxMin[iCDS] = {'max': {}, 'min': {}}
        #print(iCDS)
        for iEle in dictLoad_Inv[iCDS]: #calcolo il max ed il min per ogni sezione

            valorMAX_I = max(dictLoad_Inv[iCDS][iEle]['I'][iCDS])
            valorMAX_J = max(dictLoad_Inv[iCDS][iEle]['J'][iCDS])
            valorMIN_I = min(dictLoad_Inv[iCDS][iEle]['I'][iCDS])
            valorMIN_J = min(dictLoad_Inv[iCDS][iEle]['J'][iCDS])

            indexMaxI = dictLoad_Inv[iCDS][iEle]['I'][iCDS].index(valorMAX_I)
            indexMaxJ = dictLoad_Inv[iCDS][iEle]['J'][iCDS].index(valorMAX_J)
            indexMinI = dictLoad_Inv[iCDS][iEle]['I'][iCDS].index(valorMIN_I)
            indexMinJ = dictLoad_Inv[iCDS][iEle]['J'][iCDS].index(valorMIN_J)

            #max
            NI = dictLoad_Inv[iCDS][iEle]['I']['Axial'][indexMaxI]
            VyI = dictLoad_Inv[iCDS][iEle]['I']['Shear-y'][indexMaxI]
            VzI = dictLoad_Inv[iCDS][iEle]['I']['Shear-z'][indexMaxI]
            TI = dictLoad_Inv[iCDS][iEle]['I']['Torsion'][indexMaxI]
            MyI = dictLoad_Inv[iCDS][iEle]['I']['Moment-y'][indexMaxI]
            MzI = dictLoad_Inv[iCDS][iEle]['I']['Moment-z'][indexMaxI]

            NJ = dictLoad_Inv[iCDS][iEle]['J']['Axial'][indexMaxJ]
            VyJ = dictLoad_Inv[iCDS][iEle]['J']['Shear-y'][indexMaxJ]
            VzJ = dictLoad_Inv[iCDS][iEle]['J']['Shear-z'][indexMaxJ]
            TJ = dictLoad_Inv[iCDS][iEle]['J']['Torsion'][indexMaxJ]
            MyJ = dictLoad_Inv[iCDS][iEle]['J']['Moment-y'][indexMaxJ]
            MzJ = dictLoad_Inv[iCDS][iEle]['J']['Moment-z'][indexMaxJ]

            dictMaxMin[iCDS]['max'][iEle] =  {'I': {'Part': PartI, 'Axial': NI, 'Shear-y': VyI, 'Shear-z': VzI, 'Torsion': TI, 'Moment-y': MyI, 'Moment-z': MzI}, 
                                            'J': {'Part': PartJ, 'Axial': NJ, 'Shear-y': VyJ, 'Shear-z': VzJ, 'Torsion': TJ, 'Moment-y': MyJ, 'Moment-z': MzJ}}


            #min
            NI = dictLoad_Inv[iCDS][iEle]['I']['Axial'][indexMinI]
            VyI = dictLoad_Inv[iCDS][iEle]['I']['Shear-y'][indexMinI]
            VzI = dictLoad_Inv[iCDS][iEle]['I']['Shear-z'][indexMinI]
            TI = dictLoad_Inv[iCDS][iEle]['I']['Torsion'][indexMinI]
            MyI = dictLoad_Inv[iCDS][iEle]['I']['Moment-y'][indexMinI]
            MzI = dictLoad_Inv[iCDS][iEle]['I']['Moment-z'][indexMinI]

            NJ = dictLoad_Inv[iCDS][iEle]['J']['Axial'][indexMinJ]
            VyJ = dictLoad_Inv[iCDS][iEle]['J']['Shear-y'][indexMinJ]
            VzJ = dictLoad_Inv[iCDS][iEle]['J']['Shear-z'][indexMinJ]
            TJ = dictLoad_Inv[iCDS][iEle]['J']['Torsion'][indexMinJ]
            MyJ = dictLoad_Inv[iCDS][iEle]['J']['Moment-y'][indexMinJ]
            MzJ = dictLoad_Inv[iCDS][iEle]['J']['Moment-z'][indexMinJ]

            dictMaxMin[iCDS]['min'][iEle] = {'I': {'Part': PartI, 'Axial': NI, 'Shear-y': VyI, 'Shear-z': VzI, 'Torsion': TI, 'Moment-y': MyI, 'Moment-z': MzI}, 
                                            'J': {'Part': PartJ, 'Axial': NJ, 'Shear-y': VyJ, 'Shear-z': VzJ, 'Torsion': TJ, 'Moment-y': MyJ, 'Moment-z': MzJ}}


    return dictLoad_Inv, dictMaxMin


import pandas as pd

def ModelConci_AddSection(df_gruppi, df_secEC4):
    """
    Associa le proprietà della sezione (lette dall'Excel) ai gruppi del modello.
    - df_gruppi: DataFrame contenente i gruppi (es. 'GroupName', 'Node_I', 'Node_K', 'Node_J', ...)
    - df_secEC4: DataFrame con le proprietà delle sezioni letto in Streamlit
    """
    
    # Se il DataFrame df_secEC4 ha l'ID della sezione in una colonna (es. 'Sections') 
    # e non come indice, lo impostiamo come indice per facilitare la ricerca.
    if 'Sections' in df_secEC4.columns:
        df_secEC4 = df_secEC4.set_index('Sections')
    elif df_secEC4.index.name != 'Sections' and df_secEC4.columns[0] == 'Sections':
        # Caso fallback se la prima colonna è 'Sections'
        df_secEC4 = df_secEC4.set_index(df_secEC4.columns[0])

    # Convertiamo il dataframe delle sezioni in un dizionario {ID_sezione: {param1: val1, param2: val2, ...}}
    dictSection = df_secEC4.to_dict(orient='index')
    
    dictModelConci = {}
    
    # Otteniamo tutti i nomi dei gruppi unici dal DataFrame del nostro modello ridotto
    gruppi_univoci = df_gruppi['GroupName'].unique()
    
    for group_name in gruppi_univoci:
        # Estraiamo l'ID della sezione dal nome del gruppo
        # Es: 'group1_1' -> '1_1' -> '1' -> int(1)
        prop_str = group_name.replace('group', '').split('_')[0]
        prop_id = int(prop_str)
        
        # Popoliamo il dizionario per questa proprietà se non lo abbiamo già fatto
        if prop_id not in dictModelConci:
            if prop_id in dictSection:
                dictModelConci[prop_id] = {'section': dictSection[prop_id]}
            else:
                # Se l'utente ha inserito una sezione nel modello FEM che non è presente nel foglio Excel
                dictModelConci[prop_id] = {'section': None}
                print(f"Attenzione: La sezione con ID {prop_id} (gruppo {group_name}) non è presente nel foglio Excel delle sezioni.")
                
    return dictModelConci

def wPontiEC4_release():
    txt = '* PONTI EC4, Release: 3.3.1\n*\n*\n*\n*\n*\n+\n'
    return txt

def wPontiEC4_nConci(Model_conci):

    nConci = len(list(Model_conci.keys())) #numero conci
    txt = 'NUMERO DI CONCI\n{}\n'.format(nConci)
    return txt

def wPontiEC4_defaultMaterial():

    txt = ["MATERIALI",
            "Cls (dati comuni)",
            "fck;gammac;RH;tipo;tipoInerti;ts;t;Aesposta;uesposto;Eacc;vacc;Not_Used;alfaTerm;fcteff;chkTopBottom;wk;IsCoeffOmogInputDir",
            "35;1,5;75;N;Quarziti;0;36500;0;0;210000;0,3;0,2;1E-05;0;False;0;False",
            "Cls (Permanenti)",
            "t0(gg); psi; nE0_inpdir; nEtt0_inpdir; gammaQThermal",
            "0;1,1;0;0;1,5",
            "Cls (Ritiro)",
            "t0(gg);psi ;nE0_inpdir; nEtt0_inpdir; gammaQShrinkage",
            "0;0,55;0;0;1",
            "Cls (Deformazioni imposte, cedimenti vincolari)",
            "t0(gg);psi; nE0_inpdir; nEtt0_inpdir",
            "0;1,5;0;0",
            "Acciaio da carpenteria",
            "E(N/mm^2);v;gammaM0;gammaM1;gammaM2;gammaMser",
            "210000;0,3;1,05;1,1;1.25;1",
            "SteelGrade ;fu_l40;fu_u40;fy_l40;fy_u40",
            "EN 10025-2 S355;510;470;355;335",
            "GammaFf;GammaMF;IdGammaMfs",
            "1;1,35;0",
            "Lambda2;Lambda3;Lambda4",
            "0;0;0",
            "Acciaio pioli",
            "fu(N/mm^2) ; Gammav; ks; DeltaTauC (N/mm^2); DeltaSigmaC (N/mm^2)",
            "0;1,25;0,6;90;80",
            "GammaFf;GammaMFs;IdGammaMfs=1",
            "1;1;1",
            "LambdaV1;LambdaV2;LambdaV3;LambdaV4",
            "1,55;0;0;0",
            "Dati relativi al calcolo di LambdaV2 e Lambda2",
            "2;0;2;0;0;0;0;0",
            "Dati relativi al calcolo di LambdaV3 e Lambda3",
            "100",
            "Acciaio ordinario",
            "inp_E(N/mm^2) ; inp_fyk(N/mm^2); inp_gammas ; inp_Epsuk ; inp_kftfyk; DeltaSigmaRsk",
            "0;450;1,15;0;0;162,5",
            "GammaFf;GammaMF;IdGammaMfs=1",
            "1;1,15;1",
            "Lambda2;Lambda3;Lambda4;FiFat",
            "0;0;0;0",
            "RITIRO",
            "IsCalcAuto; DefImp(<0)",
            "True;0",
            "VAR. TERMICA",
            "Delta T",
            "10",
            "OPZIONI",
            "NumMaxIt;ErroreIt",
            "5;3"]
    
    txt2 = ""
    for i in txt:
        txt2 = txt2 + i + '\n' 

    return txt2


def wPontiEC4_Section(df_coord, dictModel_conci):
    """
    Scrive le sezioni del modello basandosi sui 3 nodi I, K, J di ogni gruppo.
    """
    txt = 'SEZIONI\n'
    
    # Iteriamo su ogni riga del DataFrame delle coordinate dei gruppi (che rappresenta un "concio")
    for id, row in df_coord.iterrows():
        group_name = row['GroupName']
        
        # Estraiamo l'ID della sezione (Property) dal nome del gruppo (es: 'group1_1' -> 1)
        # Assumiamo che il dizionario dictModel_conci abbia come chiavi l'ID della sezione
        prop_id = int(group_name.split('_')[0].replace('group', ''))
        
        # Se non trovi la proprietà nel dizionario, passa al successivo o gestisci l'errore
        if prop_id not in dictModel_conci:
            continue
            
        section_data = dictModel_conci[prop_id]['section']

        txt += f'ConcioId={id}{"-"*71}\n'
        
        txt += 'NomeConcio ; StringaElencohmet ; binf ; tinf ; bsup ; tsup ; twr ; hcop ; b1 ; IsCalcCoppella ; StringaElencoAscisse ; IsAlwaysTfCl1 ; alfaw ; IsBottomFlBox ; .RhoFlInfBox ; .IsTopFlangeDw40 ; .IsBottomFlangeDw40 ; IsApplyAdvOptTopFlange ; .IsApplyAdvOptBottomFlange\n'
        
        nome_concio = f'Concio_{group_name}'
        hs = section_data['hs (mm)']
        binf = section_data['binf (mm)']
        tinf = section_data['tinf (mm)']
        bsup = section_data['bsup (mm)']
        tsup = section_data['tsup (mm)']
        twr = section_data['tw (mm)']
        hcop = section_data['hcop (mm)']
        b1 = section_data['b1 (mm)']

        # Ora le ascisse (X) e le sezioni sono solo 3 per gruppo: I, K, J
        x_i = int(row['X_I'])
        x_k = int(row['X_K'])
        x_j = int(row['X_J'])
        X_ascisse = f"{x_i},{x_k},{x_j}"
        
        # Nomi delle 3 sezioni per questo gruppo
        NameSezioni = f"{group_name}_I,{group_name}_K,{group_name}_J"
            
        txt += f'{nome_concio};{hs};{binf};{tinf};{bsup};{tsup};{twr};{hcop};{b1};False;{X_ascisse};True;0;False;1;False;False;False;False\n'

        txt += 'StringaElencobcls ; tcls ; csup ; cinf ; pbsup ; pbinf ; Fisup ; Fiinf; StringaElencobclsSX\n'
        tcls = section_data['tcls (mm)']
        bcls = section_data['bcls (mm)']
        csup = section_data['csup (mm)']
        cinf = section_data['cinf (mm)']
        pbsup = section_data['pbsup (mm)']
        pbinf = section_data['pbinf (mm)']
        Fisup = section_data['Fisup (mm)']
        Fiinf = section_data['Fiinf (mm)']
        bsx = section_data['bsx (mm)']
        
        txt += f'{bcls};{tcls};{csup};{cinf};{pbsup};{pbinf};{Fisup};{Fiinf};{bsx}\n'

        txt += 'a ; IsRigidEndPost; h1 ; h2 ; bsldx ; tbsldx ; hsldx ; thsldx ; bslsx ; tbslsx ; hslsx ; thslsx ; IsWebSldx_NoStiff; IsWebSldx_R ; IsWebSldx_L ; IsWebSldx_T ; IsWebSlsx_NoStiff ; IsWebSlsx_R ; IsWebSlsx_L ; IsWebSlsx_T; .IsVertStiffToCheck ;.VertStiff_b1 ; .VertStiff_t1 ; .VertStiff_b2 ; .VertStiff_t2 ; .VertStiff_IsDouble ; .VertStiff_NoStiff ; .VertStiff_IsR ; .VertStiff_IsL ; .VertStiff_IsT; VertStiff_Nstex\n'
        txt += '2275;False;0;0;0;0;0;0;0;0;0;0;True;False;False;False;True;False;False;False;False;0;0;0;0;False;True;False;False;False;0\n'

        txt += 'dpioli(mm);hpioli(mm);Npioli(N/m);L(m);FxElSoletta(N);Lambda1Mom;LuceLambda1Mom;IsMidspanxLambda1;IsSupportxLambda1;Lambda1Shear;LuceLambda1Shear\n'
        dpioli = section_data['d pioli (mm)']
        hpioli = section_data['h pioli (mm)']
        npioli = section_data['n pioli (/m)']
        
        txt += f'{dpioli};{hpioli};{npioli};0;0;1;34,25;False;True;0;0\n'

        txt2 = [
            "Verifiche a fatica piattabande ed anima",
            ".DsigmaRs_Psup;DsigmaRs_Pinf;DTauRs_Web",
            "0;0;0",
            "Verifiche a fatica giunzioni piattabanda-piattabanda",
            "DsigmaRsk_PPsup;tks1_PPsup;tks2_PPsup;eks_PPsup",
            "0;30;0;0",
            "DsigmaRsk_PPinf;tks1_PPinf;tks2_PPinf;eks_PPinf",
            "0;20;0;0",
            "Verifiche a fatica composizione anima-piattabande",
            "DsigmaRs_WPsup;DsigmaRs_WPinf",
            "0;0",
            "Verifiche a fatica stiffners verticali e longitudinali",
            "DsigmaRs_VStiffW;DsigmaRs_VStiffPsup;DsigmaRs_VStiffPinf",
            "0;0;0",
            "DsigmaRs_LStiffW",
            "0",
            "Verifiche a fatica armature",
            "LambdaS1;FattTrafficoArm",
            "0;0"
        ]
        for line in txt2:
            txt += line + '\n'

        ## SEZIONI PER CONCIO
        # Ora il numero di sezioni nel concio è esattamente 3 (I, K, J)
        txt += f'Numero di sezioni nel concio\n3\n'
        txt += f'Elenco Sezioni (es. Sez1,Sez2,..)\n{NameSezioni}\n'

        txt2_avanzate = [
            'Opzioni avanzate flange',
            'False;False;False;True;0;500;False;0;0;0;0;0;0;0;0;0;0',
            '1', '1',
            'False;False;False;True;0;500;False;0;0;0;0;0;0;0;0;0;0',
            '1', '1',
            'False;False;False;True;0;250;False;0;0;0;0;0;0;0;0;0;0',
            '1', '1',
            'False;False;False;True;0;250;False;0;0;0;0;0;0;0;0;0;0',
            '1', '1'
        ]
        for line in txt2_avanzate:
            txt += line + '\n'

        # Blocco delle combinazioni standard
        txt_comb = """Comb. SLU fondamentale di Momento Massimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLU fondamentale di Momento Minimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLU fondamentale di Taglio Massimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLU fondamentale di Taglio Minimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLS caratteristica di Momento Massimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLS caratteristica di Momento Minimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLS caratteristica di Taglio Massimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLS caratteristica di Taglio Minimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLS frequente di Momento Massimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLS frequente di Momento Minimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLS frequente di Taglio Massimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SLS frequente di Taglio Minimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3b;0;0;0;0
Comb. SL fatica di Momento Massimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3bMax;0;0;0;0
F3bMin;0;0;0;0
Comb. SL fatica di Momento Minimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3bMax;0;0;0;0
F3bMin;0;0;0;0
Comb. SL fatica di Taglio Massimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3bMax;0;0;0;0
F3bMin;0;0;0;0
Comb. SL fatica di Taglio Minimo
Fase;N;V;M;T
F1;0;0;0;0
F2a;0;0;0;0
F2b;0;0;0;0;0
F2c;0;0;0;0
F3a;0;0;0;0;0
F3bMax;0;0;0;0
F3bMin;0;0;0;0
OPZIONI REPORT
IsCarattGenerali;IsDominio;IsPreclassificazione;IsSollClassFlexPlast_Mmax;IsSollClassFlexPlast_Mmin;IsSollClassFlexPlast_Vmax;IsSollClassFlexPlast_Vmax;IsSollClassFlexPlast_Tmax;IsTensioniLorde_Mmax;IsTensioniLorde_Mmin;IsTensioniLorde_Vmax;IsTensioniLorde_Tmax;IsCarattEtensioniEff_Mmax;IsCarattEtensioniEff_Mmin;IsCarattEtensioniEff_Vmax;IsCarattEtensioniEff_Tmax;IsTaglio_Mmax;IsTaglio_Mmin;IsTaglio_Vmax;IsTaglio_Tmax;IsPioliCalcElastico_Mmax;IsPioliCalcElastico_Mmin;IsPioliCalcElastico_Vmax;IsPioliCalcElastico_Tmax;IsPioliCalcPlast;IsPioliRitVarTerm
False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;False;True;False;False;False;False;False;False;False;False;False
"""
        # Ripetiamo il blocco delle combinazioni per ciascuna delle 3 sezioni
        for idSection in range(0, 3):
            txt += f'Sollecitazioni ConncioId={id}   SezioneId={idSection}\n'
            txt += txt_comb

    return txt

def wPontiEC4_final():

    #Finale
    txt2 = ["VARIATION DI BEFF",
            "0",
            "BOLTED CONNECTIONS",
            "Materials",
            "Classe",
            " ",
            "IsTaglioParteFilet",
            "True",
            "IsAnom",
            "False",
            "ftpl40",
            "0",
            "ftpu40",
            "0",
            "fypl40",
            "0",
            "fypu40",
            "0",
            "ks",
            "0",
            "mu",
            "0",
            "GammaM2",
            "0",
            "GammaM3SLU",
            "0",
            "GammaM3SLE",
            "0",
            "IdCategory",
            "0",
            "Fatigue details",
            "DtauBolt",
            "100",
            "DSigmaPlate",
            "50",
            "IsGrossPlate",
            "False",
            "IsNetPlate",
            "True",
            "DSigmaSec",
            "90",
            "IsGrossSection",
            "False",
            "IsNetSection",
            "True",
            "Connections",
            "Number of connections",
            "0",
            "OPZIONI REPORT CONNESSIONI",
            "IsMmax;IsMmax;IsMmax;IsMmax",
            "VARIATION DI hmet",
            "0"]

    txt = ''
    for i in txt2:
        txt +=i + '\n'
    
    return txt

def wPontiEC4_Model(df_coord, dictModel_conci):
    """
    Funzione principale che assembla le varie parti del file e lo salva.
    """
    txtModel = []

    # release
    txtModel.append(wPontiEC4_release())

    # numero conci
    txtModel.append(wPontiEC4_nConci(df_coord))

    # materiali
    txtModel.append(wPontiEC4_defaultMaterial())

    # sezioni
    txtModel.append(wPontiEC4_Section(df_coord, dictModel_conci))

    # parte finale
    txtModel.append(wPontiEC4_final())

    # # SCRITTURA DEL FILE .bak
    # file_path = os.path.join(pathFile, nameFile + ".bak")
    # with open(file_path, 'w') as f:
    #     for block in txtModel:
    #         f.write(block)
    # Uniamo la lista in un'unica stringa di testo
    testo_finale = "".join(txtModel)
    
    return testo_finale


def comb_PontiEC4(Model, pathInput, pathOut, NameFile):
    #path: della cartella che contiene i file excel

    #Cerco le sollecitazioni
    directory = os.path.join(pathInput)
    for root,dirs,files in os.walk(directory):
        fileList = files

    #### G1-Permanenti - Fase 1
    if ("01_Permanenti.xlsx" in fileList):
        print("File 01_Permanenti.xlsx Exists")
        dictLoad_g1 = importOneLoad_MIDAS(os.path.join(pathInput, "01_Permanenti.xlsx"))
    else:
        print("File 01_Permanenti.xlsx No Exists")

    #### G2-Permanenti - Fase 2a
    if ("02_Portati.xlsx" in fileList):
        print("File 02_Portati.xlsx Exists")
        dictLoad_g2 = importOneLoad_MIDAS(os.path.join(pathInput, "02_Portati.xlsx"))
    else:
        print("File 02_Portati.xlsx No Exists")

    #### R-Ritiro - Fase 2b
    if ("04_Ritiro.xlsx" in fileList):
        print("File 04_Ritiro.xlsx Exists")
        dictLoad_R = importOneLoad_MIDAS(os.path.join(pathInput, "04_Ritiro.xlsx"))
    else:
        print("File 04_Ritiro.xlsx No Exists")

    #### T - Temperatura - Fase 3a
    if ("05_Temperatura.xlsx" in fileList):
        print("File 05_Temperatura.xlsx Exists")
        dictLoad_temp = importMultiLoad2_MIDAS(os.path.join(pathInput, "05_Temperatura.xlsx"))
        dictLoad_Termica, dictMaxMin_Temp = inviluppoCDS_Static(dictLoad_temp)
    else:
        print("File 05_Temperatura.xlsx No Exists")

    #### MQ - Mobili Tandem 
    if ("03_Mobili_TS.xlsx" in fileList):
        print("File 03_Mobili_TS Exists")
        dictLoad_ts = importMultiLoad_MIDAS(os.path.join(pathInput, "03_Mobili_TS.xlsx"))
        dictLoad_TS, dictMaxMin_TS = inviluppoCDS_MoveLoad(dictLoad_ts)
    else:
        print("File 03_Mobili_TS.xlsx No Exists")


    #### MQ - Mobili distribuiti
    if ("03_Mobili_TS.xlsx" in fileList):
        print("File 03_Mobili_UDL.xlsx Exists")
        dictLoad_udl = importMultiLoad_MIDAS(os.path.join(pathInput, "03_Mobili_UDL.xlsx"))
        dictLoad_UDL, dictMaxMin_UDL = inviluppoCDS_MoveLoad(dictLoad_udl)
        #dictLoad_UDL = inviluppoCDS_MoveLoad(dictLoad_udl)
    else:
        print("File 03_Mobili_UDL.xlsx No Exists")

    #### Mf - Fatica
    if ("06_Fatica.xlsx" in fileList):
        print("File 06_Fatica.xlsx Exists")
        dictLoad_fatica = importMultiLoad_MIDAS(os.path.join(pathInput, "06_Fatica.xlsx"))
        dictLoad_Fat, dictMaxMin_Fas = inviluppoCDS_MoveLoad(dictLoad_fatica)
        #dictLoad_Fat = inviluppoCDS_MoveLoad(dictLoad_fatica)
    else:
        print("File 06_fatica.xlsx No Exists")
    
    # --------------------------------------------------------------------------------------------------------- 

    # Cretae a xlsx file per Ponti EC4
    xlsx_File = xlsxwriter.Workbook(NameFile + '.xlsx')

    cell_format1 = xlsx_File.add_format({'bold': True, 'font_color': '#000000', 'bg_color': '#EAEAEA'})
    cell_format2 = xlsx_File.add_format({'bold': True, 'font_color': '#000000', 'bg_color': '#808080'})
    cell_format3 = xlsx_File.add_format({'bold': False, 'font_color': '#000000', 'bg_color': '#FFFF99'})

    cell_format4 = xlsx_File.add_format({'bold': False, 'font_color': '#000000', 'bg_color': '#FFCC66'})
    cell_format5 = xlsx_File.add_format({'bold': False, 'font_color': '#000000', 'bg_color': '#FF66CC'})
    cell_format6 = xlsx_File.add_format({'bold': False, 'font_color': '#000000', 'bg_color': '#9999FF'})
    cell_format7 = xlsx_File.add_format({'bold': False, 'font_color': '#000000', 'bg_color': '#66CCFF'})
    cell_format8 = xlsx_File.add_format({'bold': False, 'font_color': '#000000', 'bg_color': '#99FFCC'})

    #worksheet
    sheetULS = xlsx_File.add_worksheet('ULS Fundamental')
    sheetSLSCar = xlsx_File.add_worksheet('SLS Characteristic')
    sheetSLSFreq = xlsx_File.add_worksheet('SLS Frequent')
    sheetFatigue = xlsx_File.add_worksheet('Fatigue')
    sheetModel = xlsx_File.add_worksheet('Model Info')

    listSheet = [sheetULS, sheetSLSCar, sheetSLSFreq, sheetFatigue]

    ## create model info sheet
    title = ['Element',	'Node',	'Section',	'Cracked',	'X(m)',	'Y(m)',	'Z(m)']
    for i, item in enumerate(title):
        sheetModel.write(0, i, item) #write intestazione


    

    title1 = ['Ponti EC4',	' ',	' ',	' ',	' ',	'N',	' ',	'V',	'T',	'M',	' ',	' ']
    title2 = ['Section',	'Element',	'GaussPt',	'Component',	'Phase',	'Fx',	'Fy',	'Fz',	'Mx',	'My',	'Mz',	'g y']
    for i, item in enumerate(title2):
        sheetULS.write(0, i, title1[i], cell_format2) #write intestazione
        sheetULS.write(1, i, item, cell_format1) #write intestazione
        sheetSLSCar.write(0, i, title1[i], cell_format2) #write intestazione
        sheetSLSCar.write(1, i, item, cell_format1) #write intestazione
        sheetSLSFreq.write(0, i, title1[i], cell_format2) #write intestazione
        sheetSLSFreq.write(1, i, item, cell_format1) #write intestazione
        sheetFatigue.write(0, i, title1[i], cell_format2) #write intestazione
        sheetFatigue.write(1, i, item, cell_format1) #write intestazione

    column = 1
    row2 = 1

    for i in Model['Element']:
        if i >= 0: # cosi escludo i nan
            ele = int(i)
            concio = int(Model['Element'][i]['Property'])
            nomeConcio = 'Concio' + str(concio)

            node = ['Node1', 'Node2']
            for j, inode in enumerate(['I', 'J']):
                nome = nomeConcio + '_' + str(ele) + inode

                #Model create
                sheetModel.write(row2, 2, nome)
                sheetModel.write(row2, 3, 'FALSO')
                X = Model['Point'][Model['Element'][i][node[j]]]['X']
                sheetModel.write(row2, 4, X)
                row2 += 1

    # WRITE CDS
    cdsName = ['Fz (Max)', 'Fz (Min)', 'My (Max)', 'My (Min)']
    phiComb = [1.20, 1.0, 1.0, 1.0]
    for ilist, iSheet in enumerate(listSheet):
        row1 = 2
        for icds in cdsName:
            for i in Model['Element']:
                if i >= 0:
                    ele = int(i)
                    concio = int(Model['Element'][i]['Property'])
                    nomeConcio = 'Concio' + str(concio)

                    node = ['Node1', 'Node2']

                    for j, inode in enumerate(['I', 'J']):
                        nome = nomeConcio + '_' + str(ele) + inode

                        # per le combinazioni di carico
                        if iSheet == sheetULS: #SLU fondamentale
                            cG1, cG2, cTS, cUDL, cR, cT = 1.35, 1.5, 1.35, 1.35, 1.2, 1.5*0.6 #coefficienti moltiplicativi
                            #γG1 · G1 + γG2 · G2 + γP · P + γQ1 · Qk1+ γQ2 · ψ02 · Qk2  +γQ3 · ψ03 · Qk3 + …
                        elif iSheet == sheetSLSCar: #SLE caratteristica (RARA)
                            cG1, cG2, cTS, cUDL, cR, cT = 1.0, 1.0, 1.0, 1.0, 1, 0.6 #coefficienti moltiplicativi
                            #G1 + G2 + P + Qk1 + ψ02 · Qk2 + ψ03 · Qk3+ …
                        elif iSheet == sheetSLSFreq: #SLE frequente
                            cG1, cG2, cTS, cUDL, cR, cT = 1.0, 1.0, 0.75, 0.40, 1.0, 0.6 #coefficienti moltiplicativi
                            #G1 + G2 + P + ψ11 · Qk1 + ψ22 · Qk2 + ψ23 · Qk3 + …
                        elif iSheet == sheetFatigue: #SLE fatica
                            cG1, cG2, cTS, cUDL, cR, cT = 1.0, 1.0, 1.0, 0.0, 1.0, 0.6 #coefficienti moltiplicativi
                            #G1 + G2 + P + ψ11 · Qk1 + ψ22 · Qk2 + ψ23 · Qk3 + …

                        ## CDS
                        #FASE 1 
                        Fx_f1 = dictLoad_g1[i][inode]['Axial']*cG1*-1000 #N in Newton
                        Fy_f1 = dictLoad_g1[i][inode]['Shear-y']*cG1*-1000
                        Fz_f1 = dictLoad_g1[i][inode]['Shear-z']*cG1*-1000 #V in Newton
                        Mx_f1 = dictLoad_g1[i][inode]['Torsion']*cG1*-1000 #T in Newton
                        My_f1 = dictLoad_g1[i][inode]['Moment-y']*cG1*-1000 #M in Newton
                        Mz_f1 = dictLoad_g1[i][inode]['Moment-z']*cG1*-1000

                        #FASE 2a 
                        Fx_f2a = dictLoad_g2[i][inode]['Axial']*cG2*-1000 #N in Newton
                        Fy_f2a = dictLoad_g2[i][inode]['Shear-y']*cG2*-1000
                        Fz_f2a = dictLoad_g2[i][inode]['Shear-z']*cG2*-1000 #V in Newton
                        Mx_f2a = dictLoad_g2[i][inode]['Torsion']*cG2*-1000 #T in Newton
                        My_f2a = dictLoad_g2[i][inode]['Moment-y']*cG2*-1000 #M in Newton
                        Mz_f2a = dictLoad_g2[i][inode]['Moment-z']*cG2*-1000

                        #FASE 2b
                        Fx_f2b = dictLoad_R[i][inode]['Axial']*cR*-1000 #N in Newton
                        Fy_f2b = dictLoad_R[i][inode]['Shear-y']*cR*-1000
                        Fz_f2b = dictLoad_R[i][inode]['Shear-z']*cR*-1000 #V in Newton
                        Mx_f2b = dictLoad_R[i][inode]['Torsion']*cR*-1000 #T in Newton
                        My_f2b = dictLoad_R[i][inode]['Moment-y']*cR*-1000 #M in Newton
                        Mz_f2b = dictLoad_R[i][inode]['Moment-z']*cR*-1000

                        #FASE 3a 
                        if icds == 'Fz (Max)':
                            cdsPontiEC4 = 'Shear-z'
                            valor = 'max'

                        elif icds == 'Fz (Min)':
                            cdsPontiEC4 = 'Shear-z'
                            valor = 'min'

                        elif icds == 'My (Max)':
                            cdsPontiEC4 = 'Moment-y'
                            valor = 'max'

                        elif icds == 'My (Min)':
                            cdsPontiEC4 = 'Moment-y'
                            valor = 'min'

                        Fx_f3a = dictMaxMin_Temp[cdsPontiEC4][valor][i][inode]['Axial']*cT*-1000 #N in Newton
                        Fy_f3a = dictMaxMin_Temp[cdsPontiEC4][valor][i][inode]['Shear-y']*cT*-1000
                        Fz_f3a = dictMaxMin_Temp[cdsPontiEC4][valor][i][inode]['Shear-z']*cT*-1000 #V in Newton
                        Mx_f3a = dictMaxMin_Temp[cdsPontiEC4][valor][i][inode]['Torsion']*cT*-1000 #T in Newton
                        My_f3a = dictMaxMin_Temp[cdsPontiEC4][valor][i][inode]['Moment-y']*cT*-1000 #M in Newton
                        Mz_f3a = dictMaxMin_Temp[cdsPontiEC4][valor][i][inode]['Moment-z']*cT*-1000

                        #FASE 3b
                        
                        if icds == 'Fz (Max)':
                            cdsPontiEC4 = 'Shear-z'
                            valor = 'max'

                        elif icds == 'Fz (Min)':
                            cdsPontiEC4 = 'Shear-z'
                            valor = 'min'

                        elif icds == 'My (Max)':
                            cdsPontiEC4 = 'Moment-y'
                            valor = 'max'

                        elif icds == 'My (Min)':
                            cdsPontiEC4 = 'Moment-y'
                            valor = 'min'
                        
                        if iSheet == sheetFatigue:
                            Fx_f3b = (dictMaxMin_Fas[cdsPontiEC4][valor][i][inode]['Axial'])*cTS*-1000 #N in Newton
                            Fy_f3b = (dictMaxMin_Fas[cdsPontiEC4][valor][i][inode]['Shear-y'])*cTS*-1000
                            Fz_f3b = (dictMaxMin_Fas[cdsPontiEC4][valor][i][inode]['Shear-z'])*cTS*-1000 #V in Newton
                            Mx_f3b = (dictMaxMin_Fas[cdsPontiEC4][valor][i][inode]['Torsion'])*cTS*-1000 #T in Newton
                            My_f3b = (dictMaxMin_Fas[cdsPontiEC4][valor][i][inode]['Moment-y'])*cTS*-1000 #M in Newton
                            Mz_f3b = (dictMaxMin_Fas[cdsPontiEC4][valor][i][inode]['Moment-z'])*cTS*-1000
                        else:
                            Fx_f3b = (dictMaxMin_TS[cdsPontiEC4][valor][i][inode]['Axial']*cTS + dictMaxMin_UDL[cdsPontiEC4][valor][i][inode]['Axial']*cUDL)*-1000 #N in Newton
                            Fy_f3b = (dictMaxMin_TS[cdsPontiEC4][valor][i][inode]['Shear-y']*cTS + dictMaxMin_UDL[cdsPontiEC4][valor][i][inode]['Shear-y']*cUDL)*-1000
                            Fz_f3b = (dictMaxMin_TS[cdsPontiEC4][valor][i][inode]['Shear-z']*cTS + dictMaxMin_UDL[cdsPontiEC4][valor][i][inode]['Shear-z']*cUDL)*-1000 #V in Newton
                            Mx_f3b = (dictMaxMin_TS[cdsPontiEC4][valor][i][inode]['Torsion']*cTS + dictMaxMin_UDL[cdsPontiEC4][valor][i][inode]['Torsion']*cUDL)*-1000 #T in Newton
                            My_f3b = (dictMaxMin_TS[cdsPontiEC4][valor][i][inode]['Moment-y']*cTS + dictMaxMin_UDL[cdsPontiEC4][valor][i][inode]['Moment-y']*cUDL)*-1000 #M in Newton
                            Mz_f3b = (dictMaxMin_TS[cdsPontiEC4][valor][i][inode]['Moment-z']*cTS + dictMaxMin_UDL[cdsPontiEC4][valor][i][inode]['Moment-z']*cUDL)*-1000
                        

                        # WRITE EXCEL

                        #FASE 1
                        iSheet.write(row1, 0, nome, cell_format4)
                        iSheet.write(row1, 3, icds, cell_format4)
                        iSheet.write(row1, 4, 'Phase 1', cell_format4)
                        iSheet.write(row1, 5, Fx_f1, cell_format4)
                        iSheet.write(row1, 6, Fy_f1, cell_format4)
                        iSheet.write(row1, 7, Fz_f1, cell_format4)
                        iSheet.write(row1, 8, Mx_f1, cell_format4)
                        iSheet.write(row1, 9, My_f1, cell_format4)
                        iSheet.write(row1, 10, Mz_f1, cell_format4)
                        iSheet.write(row1, 11, '', cell_format4)

                        #FASE 2a
                        row1 += 1
                        iSheet.write(row1, 0, nome, cell_format5)
                        iSheet.write(row1, 3, icds, cell_format5)
                        iSheet.write(row1, 4, 'Phase 2a', cell_format5)
                        iSheet.write(row1, 5, Fx_f2a, cell_format5)
                        iSheet.write(row1, 6, Fy_f2a, cell_format5)
                        iSheet.write(row1, 7, Fz_f2a, cell_format5)
                        iSheet.write(row1, 8, Mx_f2a, cell_format5)
                        iSheet.write(row1, 9, My_f2a, cell_format5)
                        iSheet.write(row1, 10, Mz_f2a, cell_format5)
                        iSheet.write(row1, 11, '', cell_format5)

                        #FASE 2b
                        row1 += 1
                        iSheet.write(row1, 0, nome, cell_format6)
                        iSheet.write(row1, 3, icds, cell_format6)
                        iSheet.write(row1, 4, 'Phase 2b', cell_format6)
                        iSheet.write(row1, 5, Fx_f2b, cell_format6)
                        iSheet.write(row1, 6, Fy_f2b, cell_format6)
                        iSheet.write(row1, 7, Fz_f2b, cell_format6)
                        iSheet.write(row1, 8, Mx_f2b, cell_format6)
                        iSheet.write(row1, 9, My_f2b, cell_format6)
                        iSheet.write(row1, 10, Mz_f2b, cell_format6)
                        iSheet.write(row1, 11, phiComb[ilist], cell_format6)

                        #FASE 3a - termica
                        row1 += 1
                        iSheet.write(row1, 0, nome, cell_format7)
                        iSheet.write(row1, 3, icds, cell_format7)
                        iSheet.write(row1, 4, 'Phase 3a', cell_format7)
                        iSheet.write(row1, 5, Fx_f3a, cell_format7)
                        iSheet.write(row1, 6, Fy_f3a, cell_format7)
                        iSheet.write(row1, 7, Fz_f3a, cell_format7)
                        iSheet.write(row1, 8, Mx_f3a, cell_format7)
                        iSheet.write(row1, 9, My_f3a, cell_format7)
                        iSheet.write(row1, 10, Mz_f3a, cell_format7)
                        iSheet.write(row1, 11, '', cell_format7)

                        # #FASE 3b  - Traffico
                        row1 += 1
                        iSheet.write(row1, 0, nome, cell_format8)
                        iSheet.write(row1, 3, icds, cell_format8)
                        iSheet.write(row1, 4, 'Phase 3b', cell_format8)
                        iSheet.write(row1, 5, Fx_f3b, cell_format8)
                        iSheet.write(row1, 6, Fy_f3b, cell_format8)
                        iSheet.write(row1, 7, Fz_f3b, cell_format8)
                        iSheet.write(row1, 8, Mx_f3b, cell_format8)
                        iSheet.write(row1, 9, My_f3b, cell_format8)
                        iSheet.write(row1, 10, Mz_f3b, cell_format8)
                        iSheet.write(row1, 11, '', cell_format8)

                        


                        row1 += 1


 
    # Close the Excel file
    xlsx_File.close()
    
    return 

