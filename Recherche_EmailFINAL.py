# -*- coding: utf-8 -*-
"""
Created on Tue Sep  7 13:43:51 2021

@author: sebas
"""
import codecs
import pandas as pd
import os
import random
import re
import sys
import traceback
from docx import Document
import PyPDF2
import argparse
from striprtf import striprtf
from pdfminer.high_level import extract_text
from tkinter.filedialog import asksaveasfile
#-----------------------------------------------------------------------------
import pandas as pd
import  tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import *
from tkinter import scrolledtext
from tkinter import messagebox
from PyPDF2 import PdfFileReader, PdfFileWriter


#-------------------------------------------------------------------------
fenetre = tk.Tk()
# Ajout d'un titre à la fenêtre principale :
fenetre.title("Application Institut Regard Persan")
tk.Label(fenetre, 
         text = "ScrolledText Widget Example", 
         font = ("Times New Roman", 15), 
         background = 'green', 
         foreground = "white")
# Personnaliser la couleur de l'arrière-plan de la fenêtre principale :
fenetre.config(bg = "#87CEEB")
# Définir les dimensions par défaut la fenêtre principale :
fenetre.geometry("1520x1080")
fenetre.resizable(width=10, height=10)
# Limiter les dimensions d’affichage de la fenêtre principale :
#fenetre.maxsize(400,300)
#fenetre.minsize(160,200)
#-------------------------------------------------------------------------
regex = re.compile(r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)")

#inutil--------------------
def interface_python():
        folder_selected = filedialog.askdirectory()
        return folder_selected
    
#-------------------------------------------------------------------------
def get_file_ext(file_name):
    tokens = file_name.split(".")
    ext = None
    if len(tokens) > 0:
        ext = tokens[-1].lower()
    return ext
#------------------------------------------------------------------------

def get_text(file_name):
    file_name = os.path.abspath(file_name)
     #file_name = askopenfilename() # show an "Open" dialog box and return the path to the selected file
    print(file_name)

    _, actual_file_name = os.path.split(file_name)
    if actual_file_name.startswith("~"):
        return ""
    print(file_name)
    
    ext = get_file_ext(file_name)
    if ext is None or ext in ["txt", "rst", "text", "adoc"]:
        try:
            with codecs.open(file_name, "r", "utf-8") as f:
                return f.read()
        except Exception:
            print("File could not be read ", file_name)
            traceback.print_exc()
    elif ext == "rtf":
        try:
            with codecs.open(file_name, "r", "utf-8") as f:
                return striprtf(f.read())
        except Exception:
            print("File could not be read ", file_name)
            traceback.print_exc()
    elif ext in ["pdf"]:
        text = extract_text(file_name)
        full_text = [text]
        with open(file_name, 'rb') as f:
            reader = PyPDF2.PdfFileReader(f)
            for pageNumber in range(reader.numPages):
                page = reader.getPage(pageNumber)
                try:
                    txt = page.extractText()
                    texte4.insert(END, txt)
                    full_text.append(txt)                    
                    #print("----: ", txt)
                except Exception:
                    print("Error PDF reader ", file_name, pageNumber)
                    traceback.print_exc()
        return "\n".join(full_text)
    elif ext in ["docx"]:
        full_text = []
        try:
            doc = Document(file_name)
            for para in doc.paragraphs:
                full_text.append(para.text)
        except Exception:
            traceback.print_exc()
        return '\n'.join(full_text)
    elif ext in ["doc"]:
        if os.name == 'nt':
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.visible = False
            _ = word.Documents.Open(file_name)
            doc = word.ActiveDocument
            return doc.Range().Text
        os.system("/Applications/LibreOffice.app/Contents/MacOS/soffice  --headless --convert-to txt:Text " + file_name)
        fileX = os.path.split(file_name)[1].split(".") + ".txt"
        try:
            with codecs.open(fileX, "r", "utf-8") as f:
                return f.read()
        except Exception:
            print("File could not be read ", fileX)
            traceback.print_exc()

    else:
        print("Unknown file extension", file_name)
    return ""
#------------------------------------------------------------------------------------------------------------------
    
#lire tous les fichiers dans le dossier
'''
def list_files(path):
    files = []
    # r = root, d = directories, f = files
    for r, d, f in os.walk(path):
        for file in f:
            files.append(os.path.join(r, file))

    lst = [file for file in files]
    #lst = askopenfilename() #[file for file in files]
    return lst'''
    
#-----------------------------------------------------------------------------------------------------------------

# lire simplement le fichier selectionné 
def list_files(file_name):
    files = []
    file_name = askopenfilename()
    files.append(file_name)
    return files
#-----------------------------------------------------------------------------------------------------------------
def get_emails(fileName):
    txt = get_text(fileName)
    l1.append(set(regex.findall(txt)))
    #derniere ligne
    return set(regex.findall(txt))

#----------------------------------------------------------------------------------------------------------------
def get_files(dir, extensions):
    if extensions is None:
        return list_files(dir)
    filtered_files = []
    for file in list_files(dir):
        ext = get_file_ext(file)
        if ext in extensions:
            filtered_files.append(file)
            print("\n---------------------------\n", file)
            #print(filtered_files)
            
    return filtered_files
#---------------------------------------------------------------------------------------------------------------

l1=[]

def main(args):
    if args.file is not None:
        files = [args.file]
    else:
        files = get_files(args.dir, args.ext)
    out = sys.stdout
    if args.dst is not None:
        out = open(args.dst, mode="w")
    else:
        
#-------------------------------------------------------------------------------------------------
        filesA = [('CSV Files', '*.csv')]
    
        fileA = asksaveasfile(filetypes = filesA, defaultextension = filesA)
        out = open(str(fileA.name), mode="w")
        # out = open(folder_selected+"/ListEmails.txt", mode="w")
        
    for file in files:
        emails = get_emails(file)
        for email in emails:
            print("\n", email)
            texte3.insert(tk.END, email)
            texte3.insert(tk.END, "\n\n")

            #print("----mail------: ", email,"\n")
            l1.append(email)
            out.write(email)
            out.write("\n")
        #df1 = pd.DataFrame(l1)
    #df1.to_csv("output1.txt") 
    if args.dst is not None:
        out.close()
 #-------------------------------------------------------------------------------------------------   
def recupere():
    #messagebox.showinfo("Alerte", entrée1.get())
    result= entrée1.get()
    print(result)
    return result

#◘chemin du dossier où sera enregistré le fichier
'''
def choisierDossier():
    global folder_selected
    folder_selected = filedialog.askdirectory()
    # Insert The path.
    texte2.insert(tk.END, folder_selected)
    return folder_selected
    #print(folder_selected )
'''
      
def close_window():
    fenetre.destroy()


#text------------------------------------------------------------------------------------------
def click():   
    #out = open(".Emails", mode="w")
    #print("Hi," + texte4.get())# Textbox widget
    try:
        if texte4.get("1.0", END)=="\n":
            messagebox.showinfo("Message",'La zone de texte à gauche est vide! Veillez coller un texte')
        else:
            emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", texte4.get('1.0', 'end'))
            df= pd.DataFrame(emails)
            df.to_csv("tesEmails.csv")
            for i in range (len( emails)):
                #out.write(emails[i])
                #out.write("\n")
                texte3.insert(tk.END, emails[i])
                texte3.insert(tk.END, "\n\n")
                
                
                print (emails)
            filesB = [('CSV Files', '*.csv')]
            
            if texte3.get("1.0", END)=="\n":
                messagebox.showinfo("Message",'Aucun email trouvé dans la zone de texte à gauche')
                exit
            else:
            
                fileB = asksaveasfile(filetypes = filesB, defaultextension = filesB)
                #out = open(fileB, mode="w")
                print(fileB.name)
                
                if len(texte3.get("1.0", END))>=1:
                    #df.to_excel("Emails.xlsx")
                    print("on")
                    df.to_csv(str(fileB.name))
                   
                else:
                    #df.to_excel(folder_selected+"/Emails.xlsx")
                    df.to_excel(fileB.name)
                    
                    #out = open(folder_selected+"/Emails.txt", mode="w")
                '''for i in range (len( emails)):
                    #out.write(emails[i])
                    #out.write("\n")
                    texte3.insert(tk.END, emails[i])
                    texte3.insert(tk.END, "\n\n")
                    print (emails)'''
    except:
          print("erreur")
   
#------------------------------------------------------------------------------------------
def lancer_script():
    parser = argparse.ArgumentParser(description='Extract emails from file')
    parser.add_argument("--dir", type=str, help="Directory/Folder Name", default=".")
    parser.add_argument("--file", type=str, help="File to parse")
    parser.add_argument("--ext", type=str, nargs='*', help="File extensions")
    parser.add_argument("--dst", type=str, help="Output file name")
    args = parser.parse_args()
    main(args)
 

#------------------------------------------------------------------------------
    # consolidation de fichiers excel
def selectFolder_Files():
    mergeData_folder = filedialog.askdirectory()
     # Insert The path.
    texte1.insert(tk.END, mergeData_folder)
    #files = os.listdir(cwd) 
    df = pd.DataFrame()
    #list_extesion = '.xlsx'$
    
    files = [('CSV Files', '*.csv')]
    file = asksaveasfile(filetypes = files, defaultextension = files)
    
    
    for file in mergeData_folder:
        if file.endswith('.xlsx'):
            df = df.append(pd.read_excel(file), ignore_index=True) 
    df.to_excel(mergeData_folder+'/mergedData.xlsx')
    
    #
    
  # Ajout d'un bouton dans la fenêtre :
#bouton4 = tk.Button (fenetre, text = "Fusionner fichiers", width='25', height = 2,
#                bg='green',command=selectFolder_Files)
#bouton3.grid(row=1, column=3, padx=1, pady=10)
#bouton4.pack()

'''#-----------------------------------------------------------------#'''
#----- Création des boutons -----##
bouton1 = tk.Button(fenetre, text='Quitter', width='25', height = 2,
                       bg='red',command= close_window)
bouton1.place(relx=1, x=0, y=-1, anchor=NE);
#bouton.grid(row=3, column=3, padx=5, pady=10)
# Ajout d'un bouton dans la fenêtre :
bouton3 = tk.Button (fenetre, text = "choisir un fichier", width='25', height = 2,
                command=lancer_script)
bouton3.place(x=510, y=560)


# Ajout d'un bouton dans la fenêtre :
bouton4 = tk.Button (fenetre, text = "GO chercher email", width='25', height = 2,
                        fg="white", bg="red",command= click)
bouton4.place(x=510, y=680)
        
texte4 = tk.scrolledtext.ScrolledText(fenetre, height = 30, width = 60,    wrap='word',
    bg='#D9BDAD')
texte4.pack(side='left', fill='y', padx=5, pady=80)

texte3 = tk.scrolledtext.ScrolledText(fenetre, height = 30, width = 60)
texte3.pack(side='left', fill='y', padx=200, pady=80)
  
'''
# Create label
label1 = Label(fenetre, text = "Colez le texte ici",width= 70,bg='red')
label1.place(x=5, y=95)
'''
def choose_pdf():
      filename = filedialog.askopenfilename(
            initialdir = "/",   # for Linux and Mac users
          # initialdir = "C:/",   for windows users
            title = "Select a File",
            filetypes = (("PDF files","*.pdf*"),("all files","*.*")))
      if filename:
          return filename

def read_pdf():
    filename = choose_pdf()
    reader = PdfFileReader(filename)
    pageObj = reader.getNumPages()
    for page_count in range(pageObj):
        page = reader.getPage(page_count)
        page_data = page.extractText()
        texte4.insert(END, page_data)
        
        
def onOpen():
    print(filedialog.askopenfilename(initialdir = "/",title = "Open file",filetypes = (("Python files","*.py;*.pyw"),("All files","*.*"))))
 
def onSave():
    print(filedialog.asksaveasfilename(initialdir = "/",title = "Save as",filetypes = (("CSV files","*.csv"),("EXCEL files","*.xlsx"))))
    

def save_file():
    my_str1=t1.get("1.0",END)  # read from one text box t1
    fob=filedialog.asksaveasfile(filetypes=[('csv','*.csv'),('text file','*..xlsx')],
        defaultextension='.txt',initialdir='D:\\my_data\\my_html',
        mode='w')
    
    
def donothing():
   x = 0


def helloCallBack():
   messagebox.showinfo( "Aide Application", "Cette application permet chercher les adresses emails dans les fichiers pdf\n\n\n"
                                            "1. BOUTON #chercher email :  permet chercher les adresses emails dans la zone de texte à gauche\n\n"
                                            "S'il trouve les emails il demande le nom de fichier où seront enregistrés les emails trouvés\n\n"
                                            
                                            "2. BOUTON #choisir un fichier : permet choisir le fichier pdf qu'on souhaite extraire les adresses emails.\n\n"
                                            "S'il trouve les emails il demande le nom de fichier où seront enregistrés les emails\n\n "
                                            "S'il ne trouve pas emails, un message s'affichera.")

  
colors = ["black", "red" , "green" , "blue"]
def color_changer():
    # choose and configure random color to the label text
    fg = random.choice(colors)
    label1.config(fg = fg)
    
    # call the color_changer() method after 200 micro seconds
    label1.after(200, color_changer)
    
    # create a list of different texts
    labels=["Colez le texte ici", "Colez le texte ici"]
    # choose and configure random text to the label
    text = random.choice(labels)
    label1.config(text=text)
    
label1 = Label(fenetre,width= 29,bg='red',font=('ariel', 20,'bold'))
label1.place(x=5, y=40)
color_changer()

label2 = Label(fenetre,text = "Liste d'emails trouvés", width= 29,bg='blue',font=('ariel', 20,'bold'))
label2.place(x=710, y=40)
color_changer()

menubar = tk.Menu(fenetre)
filemenu = tk.Menu(menubar)
filemenu.add_command(label="Ouvrir", command =read_pdf)
filemenu.add_command(label="Sauvagarder",command=onSave)
filemenu.add_command(label="quitter", command= close_window)
menubar.add_cascade(label="Fichier", menu=filemenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Application...", command=helloCallBack)
menubar.add_cascade(label="Aide", menu=helpmenu)

fenetre.config(menu=menubar)


fenetre.mainloop()
    