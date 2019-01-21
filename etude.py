import pandas as pd
import numpy as np
import xlrd

fpd = pd.ExcelFile('Projet.xlsx')
pf = pd.read_excel("Projet.xlsx")

lisEl = np.array(pf[['Nom','Prenom','Adresse','Moyenne','Age','Region','Spécialité','Sexe']]).tolist()
moyenneEcole = np.median(np.array(project_file[['Nom','Prenom','Adresse','Moyenne','Age','Region','Spécialité','Sexe']])[:,3])

noteMax=0,nbFille = 0,nbGarcon = 0
elevesAyantLaMoyenne = [],elevesAyantPlusDe20ans = [],lisReg = []
sommeMoyenneParRegion = {},nbElevesParRegion = {},moyenneParRegion = {},meilleureRegion = {}

for eleve in lisEl:
    if lisEl[lisEl.index(eleve)][3] > 10.00:
        elevesAyantLaMoyenne.append(eleve)

    if lisEl[lisEl.index(eleve)][4] > 20 :
        elevesAyantPlusDe20ans.append(eleve)

    if lisEl[lisEl.index(eleve)][7] == 'M' :
        nbGarcon +=1 
    else:
        nbFille+=1

    if lisEl[lisEl.index(eleve)][5] not in lisReg : 
        lisReg .append(lisEl[lisEl.index(eleve)][5])



for reg in lisReg  :
    for eleve in lisEl:
        if lisEl[lisEl.index(eleve)][5] == reg:
            sommeMoyenneParRegion[reg] += lisEl[lisEl.index(eleve)][3]
            nbElevesParRegion[reg] +=1

for reg in lisReg  :
    sommeMoyenneParRegion[reg] = 0
    nbElevesParRegion[reg] = 0
    moyenneParRegion[reg] = 0 

for reg in lisReg  :
    moyenneParRegion[reg] = sommeMoyenneParRegion[reg] / nbElevesParRegion[region] 
    if moyenneParRegion[reg] > noteMax:
        noteMax = moyenneParRegion[reg] 
        meilleureRegion = {} 
        meilleureRegion[reg] = moyenneParRegion[reg] 

pourcentageGarcon = (nbGarcon * 100) / 50
pourcentageFille = (nbFille * 100 ) / 50

elevesAyantLaMoyenne = pd.DataFrame(np.array(elevesAyantLaMoyenne))
elevesAyantPlusDe20ans = pd.DataFrame(np.array(elevesAyantPlusDe20ans))
statistics = pd.DataFrame(np.array([[moyenneEcole,pourcentageFille,pourcentageGarcon,list(meilleureRegion.keys())[0]]]))

elevesAyantLaMoyenne.to_excel("ElevesAyantLaMoyenne.xlsx",
                header=['Nom','Prenom','Adresse','Moyenne','Age','Region','Spécialité','Sexe'],
                index=False )
elevesAyantPlusDe20ans.to_excel("ElevesAyantPlusDe20Ans.xlsx",
                header=['Nom','Prenom','Adresse','Moyenne','Age','Region','Spécialité','Sexe'],
                index=False )
statistics.to_excel("StatisticsGlobals.xlsx",
                header=['Moyenne Ecole','Pourcentage de Fille', 'Pourcentage de Garçon', 'Meilleure Region'],
                index=False)