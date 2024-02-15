import tkinter #permet de créer des interfaces graphiques (GUI)
from tkinter import *
from tkinter import filedialog #fournit des boîtes de dialogue permettant de sélectionner des fichiers ou des répertoires à partir de l'interface graphique
from paddleocr import PaddleOCR #permet d'extraire du texte à partir d'images
import cv2 #conçue pour résoudre des problèmes de vision par ordinateur.
import os  # permet la gestion des fichiers et des répertoires et la manipulation des chemins d'accès,
from paddleocr import PaddleOCR #utilise le framework PaddlePaddle pour effectuer la reconnaissance optique de caractères (OCR)
import pandas as pd
import re
import time

#Fonction de traitement des images aprés un appui sur la touche "Entrer"
def TraitementDesImages(Event):
    path=folder_path.get()
    if path=="": #si l'utilisateur appui sur la touche entrer sans founir le chemin
        label1=tkinter.Label(root, text="Vous n'avez pas donné le chemin vers le dossier") #Message à afficher
        label1.place(x=540,y=150)


    elif os.path.exists(path): #sinon, vérifier si le chemin donné existe
        root.destroy()  #fermeture de l'interface
        # Liste des noms des images dans le dossier
        images_names = os.listdir(path)

        images_names_list=[]
        images_paths_list=[]
        for image_name in images_names:
            # Vérifier si le fichier est une image (extension .png ou .jpg)
            if image_name.lower().endswith(('.png', '.jpg')):
                # Chemin complet du fichier image
                image_path = os.path.join(path, image_name)

                images_paths_list.append(image_path)
                images_names_list.append(image_name)
                #charger l'image
                image = cv2.imread(image_path)

                # Redimensionner l'image avec une nouvelle taille
                resized_image = cv2.resize(image, (800, 800))

                # Enregistrer l'image redimensionnée avec le même nom que la première
                cv2.imwrite(image_path, resized_image)

                image = cv2.imread(image_path)

        ocr = PaddleOCR(lang='en')  # Créer une instance de l'OCR

        liste = []

        for i in range(len(images_names_list)):
            start_time = time.time()#temps de début d'extraction
            tuple_liste=["Not detected" for i in range(3)]
            tuple_liste[0] = images_names_list[i]
            result = ocr.ocr(images_paths_list[i], cls=True)

            for j in range(len(result[0])):
                pattern = r'[^0-9A-F]'
                if "B0" in result[0][j][1][0]:
                    ch = re.sub(pattern, '', result[0][j][1][0])
                    adresse_mac=""
                    for k in range(0, len(ch)-2, 2):
                        pair = ch[k:k+2]
                        adresse_mac += pair+":"
                    tuple_liste[1]=adresse_mac+result[0][j][1][0][-2]+result[0][j][1][0][-1]
                if "S/N" in result[0][j][1][0]:
                    ch=""
                    if "S/N" in result[0][j][1][0][:3]:
                        ch=result[0][j][1][0][3:]
                        tuple_liste[2]=ch
                    if ":" in tuple_liste[2]:
                        tuple_liste[2]=tuple_liste[2][1:]
                    end_time=time.time()#heure de fin de traitement
                    elapsed_time=end_time-start_time #temps écoulé

                    liste.append(tuple_liste+[elapsed_time])



        data_frame = pd.DataFrame(liste, columns=["Nom", "Adresse MAC", "Numéro de série","Temps de traitement en s"])
        excel_writer = pd.ExcelWriter('fichier.xlsx', engine='xlsxwriter')
        data_frame.to_excel(excel_writer, index=False)
        excel_writer.save()
    else:
        label1=tkinter.Label(root, text="Le chemin donné n'existe pas")
        label1.place(x=540,y=150)
#création de la fenêtre principale
root=Tk()
root.geometry("1200x400")


#création d'un label
label = tkinter.Label(root, text="Placez içi le chemin vers le dossier contenant les images à traiter")
label.place(x=10,y=100)

#Zone d'insertion du chemin du dossier
folder_path = Entry(root)
folder_path.place(x = 540 , y = 100 , width = 500)

#Exécution de la fonction TraitementDesImages() après appui sur la touche "<Return>"(Entrer)
folder_path.bind("<Return>", TraitementDesImages)
#visualiser l'interface
root.mainloop()