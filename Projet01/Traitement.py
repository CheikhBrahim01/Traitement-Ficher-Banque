import tkinter as tk
from faker import Faker
from tkinter import  filedialog
import pandas as pd
import numpy as np
import os

def generate_file():
    fake = Faker()
    nom_ficher = file_entry.get()
    number_linges = int(lines_entry.get())

    NUMREF = []
    NOM = []
    ADRESS = []
    MONTANT = []
    MATRICULE = []
    CODBOUE = []
    CODGUICH = []
    COMPTE = []
    ADRGUICH = []
    NOMBOUE = []
    EMAIL = []
    RIB = []

    for _ in range(number_linges):
        NUMREF.append(fake.ean13())
        NOM.append(fake.name())
        ADRESS.append(fake.address())
        MONTANT.append(np.random.randint(150000, 200000))
        RIB.append(np.random.randint(10, 99))
        MATRICULE.append(fake.aba())
        CODBOUE.append(fake.zipcode())
        CODGUICH.append(fake.zipcode())
        COMPTE.append("012" + fake.zipcode() + "010")
        NOMBOUE.append(fake.random_digit())
        ADRGUICH.append(fake.password())
        EMAIL.append(fake.free_email())

    df = pd.DataFrame(zip(NUMREF, NOM, ADRESS, MONTANT, MATRICULE, CODBOUE, CODGUICH, COMPTE, RIB, NOMBOUE, ADRGUICH, EMAIL),
                      columns=['NUMREF', 'NOM', 'ADRESS', 'MONTANT', 'MATRICULE', 'CODBOUE', 'CODGUICH', 'COMPTE',
                               'RIB', 'NOMBOUE', 'ADRGUICH', 'EMAIL'])

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('0', 'CCP')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('1', 'BCI')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('2', 'BEA')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('3', 'BMCI')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('4', 'ORABANK')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('5', 'SGM')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('6', 'BAMIS')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('7', 'CHINGUITTY')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('8', 'BNM')

    df['NOMBOUE'] = df['NOMBOUE'].astype('str')
    df['NOMBOUE'] = df['NOMBOUE'].replace('9', 'BPM')

    df['MONTANT'] = df['MONTANT'] - 2000

    df.loc[df.NOMBOUE == "SGM", "CODGUICH"] = "00026"
    df.loc[df.NOMBOUE == "SGM", "CODBOUE"] = "00055"
    df.loc[df.NOMBOUE == "SGM", "ADRGUICH"] = "BIIMMRMRXXX"

    df.loc[df.NOMBOUE == "BEA", "CODGUICH"] = "00025"
    df.loc[df.NOMBOUE == "BEA", "CODBOUE"] = "00054"
    df.loc[df.NOMBOUE == "BEA", "ADRGUICH"] = "AMDHMRMRXXX"

    df.loc[df.NOMBOUE == "ORABANK", "CODGUICH"] = "00025"
    df.loc[df.NOMBOUE == "ORABANK", "CODBOUE"] = "00054"
    df.loc[df.NOMBOUE == "ORABANK", "ADRGUICH"] = "ORBKMRMRXXX"

    df.loc[df.NOMBOUE == "BAMIS", "CODGUICH"] = "00024"
    df.loc[df.NOMBOUE == "BAMIS", "CODBOUE"] = "00056"
    df.loc[df.NOMBOUE == "BAMIS", "ADRGUICH"] = "BAAWMRMRXXX"

    df.loc[df.NOMBOUE == "BCI", "CODGUICH"] = "00053"
    df.loc[df.NOMBOUE == "BCI", "CODBOUE"] = "00028"
    df.loc[df.NOMBOUE == "BCI", "ADRGUICH"] = "COLIMRMRXXX"

    df.loc[df.NOMBOUE == "BMCI", "CODGUICH"] = "00001"
    df.loc[df.NOMBOUE == "BMCI", "CODBOUE"] = "00006"
    df.loc[df.NOMBOUE == "BMCI", "ADRGUICH"] = "MBICMRMRXXX"

    df.loc[df.NOMBOUE == "BNM", "CODGUICH"] = "00003"
    df.loc[df.NOMBOUE == "BNM", "CODBOUE"] = "00007"
    df.loc[df.NOMBOUE == "BNM", "ADRGUICH"] = "BQNMMRMRXXX"

    df.loc[df.NOMBOUE == "CHINGUITTY", "CODGUICH"] = "00029"
    df.loc[df.NOMBOUE == "CHINGUITTY", "CODBOUE"] = "00057"
    df.loc[df.NOMBOUE == "CHINGUITTY", "ADRGUICH"] = "CZRZMRMRXXX"

    df.loc[df.NOMBOUE == "BPM", "CODGUICH"] = "00100"
    df.loc[df.NOMBOUE == "BPM", "CODBOUE"] = "00018"
    df.loc[df.NOMBOUE == "BPM", "ADRGUICH"] = "BPMAMRMRXXX"

    df.loc[df.NOMBOUE == "CCP", "CODGUICH"] = "00004"
    df.loc[df.NOMBOUE == "CCP", "CODBOUE"] = "00008"
    df.loc[df.NOMBOUE == "CCP", "ADRGUICH"] = "BCEMMRMRXXX"

    groups = df.groupby('NOMBOUE')
    banques = {}
    for name, group in groups:
        banques[name] = group

    writer = pd.ExcelWriter(nom_ficher, engine='xlsxwriter')
    for i in banques.keys():
        banques[i].to_excel(writer, sheet_name=i)
    writer.close()

    # Open the generated file
    try:
        os.startfile(nom_ficher)  # Open the file with the default application
        status_label.config(text="File generated and opened successfully!", fg="green")
    except OSError:
        status_label.config(text="Error opening the file.", fg="red")

# GUI
root = tk.Tk()
root.title("Automatic File Generator")
root.configure(background='#41B77F')
root.geometry("1500x1500")
root.iconbitmap("logo.ico")

label_title = tk.Label(root, text="Bienvenue sur l'application", font=("Courrier", 40), bg='#41B77F',fg='white')
label_title.pack()

label_title = tk.Label(root, text="Automatiser les traitement du ficher Ecxel", font=("Courrier", 20), bg='#41B77F',fg='white')
label_title.pack()

def select_file():
    filedialog.askopenfilename()

yt_butoon = tk.Button(root, text="Sélctionner un ficher", font=("Courrier", 20), bg="#4caf50", fg="white", command=select_file)
yt_butoon.place(x=525,y=250)

label_title = tk.Label(root, text="NB: Si vous selctoinez un ficher pour faire le traitement pas obligatiore de crré un ficher par FIKER", font=("Courrier", 15), bg='#41B77F',fg='white')
label_title.pack()

file_label = tk.Label(root, text="File Name:", bg='#41B77F')
file_label.place(x=500,y=400)
file_entry = tk.Entry(root)
file_entry.place(x=500,y=420)

lines_label = tk.Label(root, text="Number of Lines:", bg='#41B77F')
lines_label.place(x=700,y=400)
lines_entry = tk.Entry(root)
lines_entry.place(x=700,y=420)

generate_button = tk.Button(root, text="Generate File", font=("Courrier", 20), bg="#4caf50", fg="white", command=generate_file, relief=tk.FLAT)
generate_button.place(x=560,y=450)


root.mainloop()
