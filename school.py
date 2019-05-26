"""
    Nous allons, dans ce qui va suivre:

    - sauvegarder notre fichier excel dans une liste avec les infos qui suivents:

    - les fonctions contenant :
        -- la liste des eleves ayant obtenu la moyenne
        -- la liste des eleves de plus de 20 ans
        -- la moyenne generale de l'ecole
        -- le pourcentage de filles de l'ecole
        -- le pourcentage de garcons de l'ecole
        -- la region qui a les meilleurs eleves
"""
import numpy as np
import pandas as pd
import xlrd
directory = xlrd.open_workbook(r"C:\Users\safi\PycharmProjects\mballo\workpage\ProjetExam.xlsx")
feuille = directory.sheet_by_index(0)
cols = feuille.ncols
rows = feuille.nrows

def listeEleve(): #fonction pour sauvegarder notre fichier ProgramExam dans une liste

    liste = []
    for r in range(1, rows):
        Nom = feuille.cell_value(r, 0)
        Prenom = feuille.cell_value(r, 1)
        Adresse = feuille.cell_value(r, 2)
        Moyenne = feuille.cell_value(r, 3)
        Age = feuille.cell_value(r, 4)
        Region = feuille.cell_value(r, 5)
        Specialite = feuille.cell_value(r, 6)
        Sexe = feuille.cell_value(r, 7)
        element = [Nom, Prenom, Adresse, Moyenne, Age, Region, Specialite, Sexe]
        liste += [element]

    return liste

def Eleve_Ayant_La_Moyenne(): #fonctio retournant la liste des eleves ayant obtenu la moyenne
    liste= listeEleve()
    Eleve_Ayant_La_Moyenne = []
    for i, element in enumerate(liste):
        if(element[3] >=10):
            Nom = liste[i][0]
            Prenom = liste[i][1]
            Moyenne = liste[i][3]
            eleve = [Nom, Prenom, Moyenne]
            Eleve_Ayant_La_Moyenne += [eleve]
    return Eleve_Ayant_La_Moyenne

def Eleve_Ayant_Plus20():
    liste = listeEleve()
    Eleve_Ayant_Plus20 = []
    for i, element in enumerate(liste):
        if(element[4] >20):
            Nom = liste[i][0]
            Prenom = liste[i][1]
            Age = liste[i][4]
            eleve = [Nom, Prenom, Age]
            Eleve_Ayant_Plus20 += [eleve]
    return Eleve_Ayant_Plus20

def Moyenne_Ecole():
    liste = listeEleve()

    somme =0
    nombre= 0
    for i, element in enumerate(liste):
        somme += liste[i][3]
        nombre += 1
    return somme/nombre

def Pourcentage_Fille():
    liste= listeEleve()
    nombreEleve = 0
    nombreFille = 0
    for i, element in enumerate(liste):
        nombreEleve += 1
        if(element[7] =="F"):
            nombreFille += 1
    return int(nombreFille*100/nombreEleve)

def Pourcentage_Garcon():
    return 100 - Pourcentage_Fille()

def regoinAyantPlusForteMoyenne():
    liste = listeEleve()


"""
   A présent nous allons proceder à l'extraction et à la creation des fichiers excels dans ce qui 
   va suivre:                                                                                             
"""
import xlwt

Fichier_Excel = xlwt.Workbook()#Creation du fichier excel

#Eleves ayant obtenu la moyenne
Moyenne = Fichier_Excel.add_sheet("Sheet_Moyenne")
liste1 = Eleve_Ayant_La_Moyenne()

for i in range(len(liste1)):
    Moyenne.write(i, 0, liste1[i][1])
    Moyenne.write(i, 1, liste1[i][0])
    Moyenne.write(i, 2, liste1[i][2])

#Eleves ages de plus de 20
Age = Fichier_Excel.add_sheet("Eleve_Ayant_Plusde20")
liste2 = Eleve_Ayant_Plus20()
for i in range(len(liste2)):
    Age.write(i, 0, liste2[i][1])
    Age.write(i, 1, liste2[i][0])
    Age.write(i, 2, liste2[i][2])

#Statistique Gloabal de l'école
stat = Fichier_Excel.add_sheet("Statistique_Global")

stat.write(0,0,Moyenne_Ecole())
stat.write(0,1,Pourcentage_Fille())
stat.write(0,2,Pourcentage_Garcon())


Fichier_Excel.save('Fichier_Excel.xls')
#fonction principale
if __name__ == "__main__":
        print("voici la liste des eleves ayant obtenu la moyenne:",Eleve_Ayant_La_Moyenne())
        print("Liste des eleves agés de plus de 20 ans:",Eleve_Ayant_Plus20())
        print("La moyenne de l'ecole est:",Moyenne_Ecole())
        print("Le pourcentage de fille dans cette ecole est de:",Pourcentage_Fille())
       # print("Fichier crée: {}".format(path))
