#!/usr/bin/python
# coding: utf-8
# By Sébastien Brière - 2016
# Importations de documents dans Alfresco

import ConfigParser
import Tix
from Tkinter import *
import codecs
import csv
import glob
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from alfREST import RESTHelper
from cmislib import CmisClient
from cmislib import Folder
from cmislib import Repository
import lib.cmislibalf
import os
import os.path
import sys
import shutil
import tkFileDialog
import tkFont
import tkMessageBox
import ttk
import uuid
import json
import time
from functools import partial

import paramiko
from contextlib import closing

reload(sys)  
sys.setdefaultencoding('utf8')

class importAlf():
    
    def __init__(self):
        self.version = "1.0.1"
        self.conf = {}
        
        self.confpack = {}
        self.fileinfo = {}
        self.Logger = {}
        self.Treatement = {}
        path = os.path.dirname(os.path.realpath(__file__)).split("/")
	fpath = path.remove("lib")
        self.dir_path = "/".join(path)
        self.conf = self.get_conf()
        self.conf['open'] = ""

    # Get Global Config
    def get_conf(This):
        config = ConfigParser.ConfigParser()
        
        file = open(This.dir_path+"/importalf.conf", "rb")

        config.readfp(file)
        
        file.close()
        
        conf = {}
        try:
            conf['urltemp'] = config.get("GLOBAL", "url")
            if ( conf['urltemp'] == "" ):
                This.logger(LOG+"urltemp","Error")
        except:
            This.logger(LOG+"urltemp","Error")
            conf["urltemp"] = ""
        
        try:
            conf['user'] = config.get("GLOBAL", "user")
            if ( conf['user'] == "" ):
                This.logger(LOG+"user","Error")
        except:
            This.logger(LOG+"user","Error")
            conf["user"] = ""
        
        try:
            conf['password'] = config.get("GLOBAL", "password")
            if ( conf['password'] == "" ):
                This.logger(LOG+"password","Error")
        except:
            This.logger(LOG+"password","Error")
            conf["password"]
        
        try:
            conf['host'] = config.get("GLOBAL", "host")
            if ( conf['host'] == "" ):
                This.logger(LOG+"host","Error")
        except:
            This.logger(LOG+"host","Error")
            conf["host"] = ""
        
        try:
            conf['importdir'] = config.get("GLOBAL", "bulkimportdir")
            if ( conf['importdir'] == "" ):
                This.logger(LOG+"importdir","Error")
        except:
            This.logger(LOG+"importdir","Error")
            conf["importdir"] = ""
            
        try:
            conf['user_ssh'] = config.get("GLOBAL", "user_ssh")
            if ( conf['user_ssh'] == "" ):
                This.logger(LOG+"user_ssh","Error")
        except:
            This.logger(LOG+"user_ssh","Error")
            conf["user_ssh"] = ""
        
        try:
            conf['password_ssh'] = config.get("GLOBAL", "password_ssh")
            if ( conf['password_ssh'] == "" ):
                This.logger(LOG+"password_ssh","Error")
        except:
            This.logger(LOG+"password_ssh","Error")
            conf["password_ssh"] = "" 
            
        conf['port'] = "8080"    
        return conf

    # Get Package Config
    def get_confpack(This):
        config = ConfigParser.ConfigParser()
        file = open(This.conf['dir'] + "Conf/package.conf", "rb")
        config.readfp(file)

        This.confpack = {}
        This.confpack['name'] = int(config.get("PACKAGE", "OLDNAME"))
        This.confpack['newname'] = int(config.get("PACKAGE", "NEWNAME"))
        This.confpack['title'] = int(config.get("PACKAGE", "TITLE"))
        This.confpack['desc'] = int(config.get("PACKAGE", "DESC"))
        This.confpack['tags'] = int(config.get("PACKAGE", "TAGS"))
        This.confpack['path'] = int(config.get("PACKAGE", "DESTPATH"))
        return This.confpack

    # Logger
    def logger(This, txt, tag, silence=False):
        if ( silence == False ):
            This.Logger.insert("end", txt + "\n",tag)
            This.Logger.see("end")
            This.Treatment.update()

    # Generate and get the packageId
    def get_packageid(This):
        if (os.path.exists(This.conf['dir'] + "Conf/packageId")):
            file = open(This.conf['dir'] + "Conf/packageId", "rb")
            PKGID = file.read()
            file.close()
        else:
            PKGID = str(uuid.uuid4())
            file = open(This.conf['dir'] + "Conf/packageId", "wb")
            file.write(PKGID)
            file.close()

        return PKGID

    def UpdateGuide(This,txt):
            This.Guide.delete('1.0', END)
            This.Guide.insert("1.0", txt ,"tag-center")
    
    # Gui Composer
    
    # Taille et position de la fenêtre
    def posfen(This,fen, FENW=0, FENH=0):
        RESX = fen.winfo_screenwidth()
        if ( RESX > 1920 ):
            RESX = 1920
        RESY = fen.winfo_screenheight()
        
        if ( FENW == 0 ):
            fen.after(500,fen.update_idletasks())
            FENW = fen.winfo_reqwidth()
            FENH = fen.winfo_reqheight()
        
        POSX = (RESX - FENW) /2
        POSY = (RESY - FENH) /2

        fen.geometry(str(FENW)+"x"+str(FENH)+"+"+str(POSX)+"+"+str(POSY))
        fen.deiconify()
    
    def askYesNo(This, title, message, parent):
        return tkMessageBox.askquestion(title,message,parent=parent)
    
    def messageShow(This, title, message, parent):
        return tkMessageBox.showwarning(title,message,parent=parent)
    
    def get_fieldsCSV(This):
        csvfile = open(This.conf['dir'] + 'Conf/list.csv', "rb")
        reader = csv.reader(csvfile, delimiter=';', quotechar='"')

        fields = {}

        for rows in reader:
            i=0
            for row in rows:
                fields[i] = unicode(row,"iso8859_1")
                i=i+1
            break
            
        csvfile.close()
        
        return fields
    
    def getIndex(self,value,dicto):
        return dicto.keys()[dicto.values().index(value)] 
    
    def gen_dialog(This):
        def CommandClearLog():
            This.Logger.delete('1.0', END)
        
        def CommandGenerate():
            This.Logger.delete('1.0', END)
  
            if (os.path.exists(This.conf['dir'] + "Conf/package.conf")):
                if ( os.path.exists(This.conf['dir']+"/"+This.confpack['PKGID']) == False ):
                    os.mkdir(This.conf['dir']+"/"+This.confpack['PKGID'])
                else:
                    shutil.rmtree(This.conf['dir']+"/"+This.confpack['PKGID'])
                    os.mkdir(This.conf['dir']+"/"+This.confpack['PKGID'])
                    
                initfile=open(This.conf['dir'] + This.confpack['PKGID'] +"/Sites.metadata.properties.xml","wb")
                initfile.close()
                RESULT = This.generatepack()
                if ( RESULT == True ):
                    This.ButtonGenerate['state'] = "disabled"
                    This.ButtonUpload['state'] = "active"
                    This.Force['state'] = "active"
                    This.logger("Génération des documents réussie","Success")
                    This.UpdateGuide("Etape 3 : Importez les documents")    

        def CommandOpenPackage(path=""):
            This.Logger.delete('1.0', END)
            
            if ( path == "" ):
                # Dialog Box for choose dir package
                PackageDir = tkFileDialog.askdirectory(initialdir=This.dir_path, title="Choisir le répertoire")
                PathDir['text'] = PackageDir
                This.conf['dir'] = PathDir['text'] + "/"
            else:
                This.conf['dir'] = path

            check_package_conf = False
            check_aspects_conf = False
            check_properties_conf = False

            if (os.path.exists(This.conf['dir'] + "Conf/package.conf")):
                This.confpack = This.get_confpack()
            else:
                This.logger("\nLe fichier package.conf est introuvable\n","Error")
                
            if (os.path.exists(This.conf['dir'] + "Conf/list.csv")):    
                This.fileinfo = This.parse_csv()

                This.logger("Nombre de documents : "+str(len(This.fileinfo)),"")

                fields = This.get_fieldsCSV()

                name = fields[This.confpack['name']]
                newname = fields[This.confpack['newname']]
                title = fields[This.confpack['title']]
                description = fields[This.confpack['desc']]
                tags = fields[This.confpack['tags']]
                path = fields[This.confpack['path']]

                This.logger("\nVérification correspondance des champs :\n","")
                This.logger("  Attributs   | Champs CSV","")
                This.logger("  ____________|___________    ","")
                This.logger("  OLDNAME     | "+name,"")
                This.logger("  NEWNAME     | "+newname,"")
                This.logger("  TITLE       | "+title,"")
                This.logger("  DESC        | "+description,"")
                This.logger("  TAGS        | "+tags,"")
                This.logger("  DESTPATH    | "+path,"")

                check_package_conf = True
            else:
                This.logger("\nLe fichier list.csv est introuvable\n","Error")

            if (os.path.exists(This.conf['dir'] + "Conf/aspects.conf") and check_package_conf ):    

                This.logger("\nListe des aspects :\n","")

                file = open(This.conf['dir'] + "Conf/aspects.conf", "r")
                for aspect in file.readlines():
                    This.logger("--> "+aspect.rstrip(),"")
                file.close()
                check_aspects_conf = True
            else:
                if ( check_package_conf ):
                    This.logger("\nLe fichier aspects.conf est introuvable\n","Error")

            if (os.path.exists(This.conf['dir'] + "Conf/properties.csv") and check_package_conf ):    
                This.logger("\nListe des propriétés :\n","")

                csvfile = open(This.conf['dir'] + 'Conf/properties.csv', "rb")
                reader = csv.reader(csvfile, delimiter=';', quotechar='"')

                for row in reader:
                    name = row[0]
                    format = row[1]
                    type = row[2]
                    field= row[3]

                    if ( type == "STA" ):
                        valuedetail = " : Valeur statique '"+unicode(field, "iso8859_1")+"'"
                    else:
                        valuedetail = " : Valeur dynamique du champs CSV '"+fields[int(field)]+"'"

                    This.logger("--> : "+name+valuedetail,"")
                    check_properties_conf = True
                csvfile.close()
            else:
                if ( check_package_conf ):
                    This.logger("\nLe fichier properties.conf est introuvable\n","Error")       

            if ( check_package_conf and check_aspects_conf and check_properties_conf ):
                PKGID = This.get_packageid()
                
                This.confpack['PKGID'] = PKGID

                This.logger("\nPackage ID : "+This.confpack['PKGID'], "Success")
                This.logger(This.conf['dir'] + " : OK.", "Success")

                This.ButtonGenerate['state'] = "active"
                This.ButtonTestHost['state'] = "active"
                This.UpdateGuide("Etape 2 : Générez le package")
                This.conf['open'] = This.confpack['PKGID']
            else:
                This.logger("\n"+This.conf['dir'] + " : Erreur.", "Error")
                This.ButtonGenerate['state'] = "disabled"
                This.ButtonTestHost['state'] = "disabled"
                This.ButtonUpload['state'] = "disabled"

        def CommandUpload():
            mode = "BULKIMPORTTOOL"
            
            if ( This.var1.get() == 1 ):
                mode = "CMIS"
            
            TestHost = This.OpenHost(mode,True)
            
            if ( TestHost ):
                if (os.path.exists(This.conf['dir'] + "Conf/package.conf")):
                    if ( mode == "CMIS" ):
                        RESULT = This.upload(mode)
                    else:
                        RESULT = This.initiateBulkImport()
                    if ( RESULT == True ):
                        This.logger("Import des documents réussie","Success")
                        This.UpdateGuide("Terminé")
                        This.ButtonUpload['state'] = "active"
                        This.Force['state'] = "active"
                    else:
                        if ( RESULT == "Missing" ):
                            This.logger("Import des documents réussi, mais quelques documents manquants","Success")
                        else:
                            This.logger("Import des documents échoué","Error")
                    
        def ChangeMode():
            if ( This.var1.get() == 0 ):
                This.ButtonUpload['text'] = "Importer (mode Bulk Import)"
            else:
                This.ButtonUpload['text'] = "Mettre à jour (mode CMIS)"
        
        def CommandTestHost():
            mode = "BULKIMPORTTOOL"
            TestHost = This.OpenHost(mode, False)
        
        def CommandConfGlobale():
            def SaveConf():
                file = open(This.dir_path+"/importalf.conf", "wb")
                file.write("[GLOBAL]\n")
                file.write("url="+UrlTemp.get()+"\n")
                file.write("user="+User.get()+"\n")
                file.write("password="+Cred.get()+"\n")
                file.write("host="+Host.get()+"\n")
                file.write("bulkimportdir="+BulkImportDir.get()+"\n")
                file.write("user_ssh="+UserSSH.get()+"\n")
                file.write("password_ssh="+CredSSH.get()+"\n")
                file.close()
                This.conf = This.get_conf()
                This.conf['dir'] = PathDir['text'] + "/"
                fenconf.destroy()

            fenconf = Toplevel(fen)
            fenconf.title('Import Alfresco - Configuration générale')
            
                
            defaultfont = tkFont.Font(fenconf, size=10, family='Verdana', weight='bold')
            helpfont = tkFont.Font(fenconf, size=10, family='Verdana', slant='italic')
            buttonfont = tkFont.Font(fenconf, size=10, family='Verdana', weight='bold')
            
            LabelUrlTemp = Label(fenconf, anchor=W, text="URL Template : ", width=20, font=defaultfont)
            HelpUrlTemp = Label(fenconf, anchor=W,justify="left", text="L'URL template doit comporter la chaîne __HOST__\nExemple : http://__HOST__:8080/alfresco/api/-default-/public/cmis/versions/1.1/atom", width=70, font=helpfont)
            LabelUser = Label(fenconf, anchor=W, text="Login Alfresco : ", width=20, font=defaultfont)
            LabelCred = Label(fenconf, anchor=W, text="Password Alfresco : ", width=20, font=defaultfont)
            LabelHost = Label(fenconf, anchor=W, text="Serveur hôte Alfreco : ", width=20, font=defaultfont)
            LabelBulkImportDir = Label(fenconf, anchor=W, text="Bulk Import Directory : ", width=20, font=defaultfont)
            LabelUserSSH = Label(fenconf, anchor=W, text="Login SSH : ", width=20, font=defaultfont)
            LabelCredSSH = Label(fenconf, anchor=W, text="Password SSH : ", width=20, font=defaultfont)
            
            
            UrlTemp = Entry(fenconf,bg="white",width=60)
            UrlTemp.insert(END,This.conf['urltemp'])
            
            User = Entry(fenconf,bg="white",width=15)
            User.insert(END,This.conf['user'])
            
            Cred = Entry(fenconf,show="*",bg="white",width=15)
            Cred.insert(END,This.conf['password'])
            
            UserSSH = Entry(fenconf,bg="white",width=15)
            UserSSH.insert(END,This.conf['user_ssh'])
            
            CredSSH = Entry(fenconf,show="*",bg="white",width=15)
            CredSSH.insert(END,This.conf['password_ssh'])
            
            Host = Entry(fenconf,bg="white",width=45)
            Host.insert(END,This.conf['host'])
            
            BulkImportDir = Entry(fenconf,bg="white",width=45)
            BulkImportDir.insert(END,This.conf['importdir'])
            
            Quit = Button(fenconf, text="Quitter", command=fenconf.destroy, relief=RAISED, font=buttonfont)
            Save = Button(fenconf, text="Sauver", command=SaveConf, relief=RAISED, font=buttonfont)
            
            LabelUrlTemp.grid(row=0, column=0,padx=3,pady=1)
            UrlTemp.grid(sticky="W",row=0, column=1,padx=3,pady=1)
            HelpUrlTemp.grid(sticky="W",row=1, column=1,padx=3,pady=1)
            LabelUser.grid(row=2, column=0,padx=3,pady=1)
            User.grid(sticky="W",row=2, column=1,padx=3,pady=1)
            LabelCred.grid(row=3, column=0,padx=3,pady=1)
            Cred.grid(sticky="W",row=3, column=1,padx=3,pady=1)
            LabelHost.grid(row=4, column=0,padx=3,pady=1)
            Host.grid(sticky="W",row=4, column=1,padx=3,pady=1)
            LabelBulkImportDir.grid(row=6, column=0,padx=3,pady=1)
            BulkImportDir.grid(sticky="W",row=6, column=1,padx=3,pady=1)
            LabelUserSSH.grid(row=7, column=0,padx=3,pady=1)
            UserSSH.grid(sticky="W",row=7, column=1,padx=3,pady=1)
            LabelCredSSH.grid(row=8, column=0,padx=3,pady=1)
            CredSSH.grid(sticky="W",row=8, column=1,padx=3,pady=1)
            Save.grid(sticky="W",row=9, column=0,padx=3,pady=1)
            Quit.grid(sticky="E",row=9, column=1,padx=3,pady=1)
            
            fenconf.withdraw()
            This.posfen(fenconf)
            return
        
        def gen_listFields(box):
            fields = This.get_fieldsCSV()
            for field in fields:
                box.insert(field,fields[field])
            return
        
        def CommandCreatePackage():
            def PkgPlace():
                Place = tkFileDialog.askdirectory(parent=fenconf,title="Choisir le CSV", initialdir=This.dir_path) 
                try:
                    This.conf['dir'] = Place+"/"
                    if (os.path.exists(This.conf['dir'] + "Conf") == False):
                        os.mkdir(This.conf['dir']+"Conf")
                    if (os.path.exists(This.conf['dir'] + "Orig") == False):
                        os.mkdir(This.conf['dir']+"Orig")
                    This.get_packageid()
                    This.ButtonPkgPlace['state'] = "disabled"
                    This.ButtonImportCSV['state'] = "active"
                    
                    return True
                except Exception, e:
                    This.logger(str(e),"Error")
                    return False
            
            def ImportCSV():
                try:
                    if (os.path.exists(This.conf['dir'] + "Conf/list.csv") == True):
                        rep = This.askYesNo("Attention","Le package en cours de création/modification contient déjà un fichier list.csv.\nConfirmez vous le nouvel import ?",fenconf)
                        if ( rep == "yes" ):
                            importcsv = True
                        else:
                            importcsv = False
                    else:
                        importcsv = True

                    if ( importcsv == True ):
                        CSV = tkFileDialog.askopenfilename(parent=fenconf,title="Choisir le CSV", initialdir=This.dir_path, \
                            initialfile="", filetypes = [("Fichiers CSV","*.csv")]) 
                        shutil.copy2(CSV,This.conf['dir'] + "Conf/list.csv")
                        
                    This.ButtonImportCSV['state'] = "disabled"
                    This.ButtonConfig['state'] = "active"
                    This.ButtonAspects['state'] = "active"
                    This.ButtonProperties['state'] = "active"
                    This.ButtonImportFiles['state'] = "active"
                    This.logger("Le fichier CSV a été importé.","Success")
                    return True
                except Exception, e:
                    return False

            def ImportFiles():
                This.confpack = This.get_confpack()
                Place = tkFileDialog.askdirectory(parent=fenconf,title="Choisir le dossier contenant les documents", initialdir=This.dir_path)
                try:
                    sourceDir = Place+"/"
                    fileinfo = This.parse_csv()
                    for line in fileinfo:
                        if ( os.path.exists(sourceDir + fileinfo[line]['name']) == True):
                            shutil.copy2(sourceDir + fileinfo[line]['name'],This.conf['dir'] + "Orig/" + fileinfo[line]['name'])
                        else:
                            This.logger("Document du CSV manquant dans le répertoire","Error")
                            return False
                    This.ButtonPkgPlace['state'] = "disabled"
                    This.logger("Les documents ont été importés.","Success")
                    return True
                except Exception, e:
                    This.logger(str(e),"Error")
                    return False
            
            def Config():
                def saveConfPkg():
                    complete = True
                    if ( var['OLDNAME'].get() == "" ):
                        complete = False
                    if ( var['NEWNAME'].get() == "" ):
                        complete = False
                    if ( var['TITLE'].get() == "" ):
                        complete = False
                    if ( var['DESC'].get() == "" ):
                        complete = False
                    if ( var['TAGS'].get() == "" ):
                        complete = False
                    if ( var['DESTPATH'].get() == "" ):
                        complete = False
                    
                    if ( complete == True ):
                        fields = This.get_fieldsCSV()
                        try:
                            file = open(This.conf['dir']+"Conf/package.conf", "wb")
                            file.write("[PACKAGE]\n")
                            file.write("OLDNAME="+str(This.getIndex(var['OLDNAME'].get(),fields))+"\n")
                            file.write("NEWNAME="+str(This.getIndex(var['NEWNAME'].get(),fields))+"\n")
                            file.write("TITLE="+str(This.getIndex(var['TITLE'].get(),fields))+"\n")
                            file.write("DESC="+str(This.getIndex(var['DESC'].get(),fields))+"\n")
                            file.write("TAGS="+str(This.getIndex(var['TAGS'].get(),fields))+"\n")
                            file.write("DESTPATH="+str(This.getIndex(var['DESTPATH'].get(),fields))+"\n")
                            file.close()
                        except Exception, e:
                            This.logger(str(e),"Error")
                        fenconfpkg.destroy()
                        This.ButtonAspects['state'] = "active"
                    else:
                        This.messageShow("Attention","Vous devez renseigner tous les champs.", fenconfpkg)
                    
                def updateList(evt, name, key, values):
                    value = name.entry.get()
                    
                    already = False
                    for k in values:
                        if ( k != name and value == values[k]):
                            already=True
                            
                    if ( already == False ):
                        values[key] = value
                    else:
                        This.messageShow("Attention","Champ déjà utilisé !!", fenconfpkg)
                        var[key].set("")
                    
                fenconfpkg = Toplevel(fenconf)
                fenconfpkg.title("Configutation (package.conf)")

                LabelKeys = Label(fenconfpkg, anchor=W, text="Clefs", width=20, font=titlefont)
                LabelFields = Label(fenconfpkg, anchor=W, text="Champs (CSV)", width=20, font=titlefont)

                LabelKeys.grid(row=0, column=0,padx=3,pady=1)
                LabelFields.grid(row=0, column=1,padx=3,pady=1)

                LabelOLDNAME = Label(fenconfpkg, anchor=W, text="Fichier PDF", width=20, font=defaultfont)
                LabelNEWNAME = Label(fenconfpkg, anchor=W, text="Nom d'import", width=20, font=defaultfont)
                LabelTITLE = Label(fenconfpkg, anchor=W, text="Titre", width=20, font=defaultfont)
                LabelDESC = Label(fenconfpkg, anchor=W, text="Description", width=20, font=defaultfont)
                LabelTAGS = Label(fenconfpkg, anchor=W, text="TAGS", width=20, font=defaultfont)
                LabelDESTPATH = Label(fenconfpkg, anchor=W, text="Destination", width=20, font=defaultfont)
                
                var = {}
                values={}
                
                var['OLDNAME'] = Tix.StringVar()  
                listOLDNAME = Tix.ComboBox(fenconfpkg, editable=1, dropdown=1, variable=var['OLDNAME'])
                gen_listFields(listOLDNAME)
                listOLDNAME.slistbox.listbox.bind('<ButtonRelease-1>', partial(updateList,name=listOLDNAME,key="OLDNAME",values=values))

                var['NEWNAME'] = Tix.StringVar()  
                listNEWNAME = Tix.ComboBox(fenconfpkg, editable=1, dropdown=1, variable=var['NEWNAME'])
                gen_listFields(listNEWNAME)
                listNEWNAME.slistbox.listbox.bind('<ButtonRelease-1>', partial(updateList,name=listNEWNAME,key="NEWNAME",values=values))

                var['TITLE'] = Tix.StringVar()  
                listTITLE = Tix.ComboBox(fenconfpkg, editable=1, dropdown=1, variable=var['TITLE'])
                gen_listFields(listTITLE)
                listTITLE.slistbox.listbox.bind('<ButtonRelease-1>', partial(updateList,name=listTITLE,key="TITLE",values=values))

                var['DESC'] = Tix.StringVar()  
                listDESC = Tix.ComboBox(fenconfpkg, editable=1, dropdown=1, variable=var['DESC'])
                gen_listFields(listDESC)
                listDESC.slistbox.listbox.bind('<ButtonRelease-1>', partial(updateList,name=listDESC,key="DESC",values=values))

                var['TAGS'] = Tix.StringVar()  
                listTAGS = Tix.ComboBox(fenconfpkg, editable=1, dropdown=1, variable=var['TAGS'])
                gen_listFields(listTAGS)
                listTAGS.slistbox.listbox.bind('<ButtonRelease-1>', partial(updateList,name=listTAGS,key="TAGS",values=values))

                var['DESTPATH'] = Tix.StringVar()  
                listDESTPATH = Tix.ComboBox(fenconfpkg, editable=1, dropdown=1, variable=var['DESTPATH'])
                gen_listFields(listDESTPATH)
                listDESTPATH.slistbox.listbox.bind('<ButtonRelease-1>', partial(updateList,name=listDESTPATH,key="DESTPATH",values=values))
                
                if (os.path.exists(This.conf['dir'] + "Conf/package.conf") == True):
                    confpacktmp = This.get_confpack()
                    fields = This.get_fieldsCSV()
                    var['OLDNAME'].set(fields[confpacktmp['name']])
                    var['NEWNAME'].set(fields[confpacktmp['newname']])
                    var['TITLE'].set(fields[confpacktmp['title']])
                    var['DESC'].set(fields[confpacktmp['desc']])
                    var['TAGS'].set(fields[confpacktmp['tags']])
                    var['DESTPATH'].set(fields[confpacktmp['path']])

                LabelKeys.grid(row=0, column=0,padx=3,pady=1)
                LabelFields.grid(row=0, column=1,padx=3,pady=1)

                LabelOLDNAME.grid(row=1, column=0,padx=3,pady=1)
                LabelNEWNAME.grid(row=2, column=0,padx=3,pady=1)
                LabelTITLE.grid(row=3, column=0,padx=3,pady=1)
                LabelDESC.grid(row=4, column=0,padx=3,pady=1)
                LabelTAGS.grid(row=5, column=0,padx=3,pady=1)
                LabelDESTPATH.grid(row=6, column=0,padx=3,pady=1)

                listOLDNAME.grid(row=1, column=1,padx=3,pady=1)
                listNEWNAME.grid(row=2, column=1,padx=3,pady=1)
                listTITLE.grid(row=3, column=1,padx=3,pady=1)
                listDESC.grid(row=4, column=1,padx=3,pady=1)
                listTAGS.grid(row=5, column=1,padx=3,pady=1)
                listDESTPATH.grid(row=6, column=1,padx=3,pady=1)

                Quit = Button(fenconfpkg, text="Quitter", command=fenconfpkg.destroy, relief=RAISED, font=buttonfont)
                Save = Button(fenconfpkg, text="Sauver", command=saveConfPkg, relief=RAISED, font=buttonfont)

                Save.grid(sticky="W",row=7, column=0,padx=3,pady=1)
                Quit.grid(sticky="E",row=7, column=1,padx=3,pady=1)

                fenconfpkg.withdraw()
                This.posfen(fenconfpkg)
                return
            
            def ConfigAspects():
                def saveAspects():
                    listaspects = Aspects.get('1.0', 'end')
                    tabaspects = listaspects.split("\n")
                    try:
                        file = open(This.conf['dir']+"Conf/aspects.conf", "wb")
                        for aspect in tabaspects:
                            if ( aspect != "" ):
                                file.write(aspect+"\n")
                        file.close()
                        fenconfasp.destroy()
                        This.ButtonProperties['state'] = "active"
                    except Exception, e:
                        This.logger(str(e),"Error")
                    return
                
                fenconfasp = Toplevel(fenconf)
                fenconfasp.title("Configutation des aspects")
                
                Aspects = Text(fenconfasp,bg="white",width=30,height=10)
                
                if (os.path.exists(This.conf['dir'] + "Conf/aspects.conf") == True):
                    file = open(This.conf['dir'] + "Conf/aspects.conf", "r")
                    line = 1
                    for aspect in file.readlines():
                        if ( aspect != "\n" ):
                            Aspects.insert(str(line)+".0",aspect)
                            line=line+1
                    file.close()
                
                Aspects.grid(sticky="W",row=0, columnspan=2,padx=3,pady=1)

                Quit = Button(fenconfasp, text="Quitter", command=fenconfasp.destroy, relief=RAISED, font=buttonfont)
                Save = Button(fenconfasp, text="Sauver", command=saveAspects, relief=RAISED, font=buttonfont)

                Save.grid(sticky="W",row=1, column=0,padx=3,pady=1)
                Quit.grid(sticky="E",row=1, column=1,padx=3,pady=1)

                fenconfasp.withdraw()
                This.posfen(fenconfasp)
                return
            
            def ConfigProperties():
                def updateList(evt, name, key, word):
                    choice = name.entry.get()
                    props[key][word] = choice
                    
                def addSta():
                    add("STA")
                    
                def addDyn():
                    add("DYN")
                 
                def add(type):
                    idx = len(props) + 1
                    props[idx] = {}
                    props[idx]['type'] = type
                    props[idx]['name'] = Entry(Block,bg="white",width=15)
                    
                    props[idx]['format'] = ""
                    props[idx]['formattix'] = Tix.StringVar()  
                    props[idx]['formatlist'] = Tix.ComboBox(Block, editable=1, dropdown=1, variable=props[idx]['formattix'], width=20, listwidth=15)
                    props[idx]['formatlist'].insert(0,"TXT")
                    props[idx]['formatlist'].insert(1,"NUM")
                    props[idx]['formatlist'].insert(2,"DATE")
                    props[idx]['formatlist'].entry.config(width=8, state='readonly')
                    props[idx]['formatlist'].slistbox.listbox.bind('<ButtonRelease-1>',partial(updateList,name=props[idx]['formatlist'],key=idx,word="format"))
                    
                    if ( type == "STA" ):
                        props[idx]['value'] = Entry(Block,bg="white",width=15)
                    else:
                        props[idx]['value'] = ""
                        props[idx]['valuetix'] = Tix.StringVar()
                        props[idx]['valuelist'] = Tix.ComboBox(Block, editable=1, dropdown=1, variable=props[idx]['valuetix'], width=20, listwidth=50)
                        gen_listFields(props[idx]['valuelist'])
                        props[idx]['valuelist'].entry.config(width=15, state='readonly')
                        props[idx]['valuelist'].slistbox.listbox.bind('<ButtonRelease-1>',partial(updateList,name=props[idx]['valuelist'],key=idx,word="value"))
                    
                    props[idx]['name'].grid(sticky="W",row=idx, column=0,padx=3,pady=1)
                    props[idx]['formatlist'].grid(sticky="W",row=idx, column=1,padx=3,pady=1)
                    if ( type == "STA" ):
                        props[idx]['value'].grid(sticky="W",row=idx, column=2,padx=3,pady=1)
                    else:
                        props[idx]['valuelist'].grid(sticky="W",row=idx, column=2,padx=3,pady=1)
                    Block.update_idletasks()
                    return
                            
                def saveProperties():
                    file = open(This.conf['dir']+"Conf/properties.csv", "wb")
                    fields = This.get_fieldsCSV()
                    for idx in props:
                        name = props[idx]['name'].get()
                        format = props[idx]['format']
                        type = props[idx]['type']
                        if ( type == "STA" ):
                            value = props[idx]['value'].get().decode("utf-8").encode("iso8859_1")
                        else:
                            value = str(This.getIndex(props[idx]['value'],fields))
                        file.write(name+";"+format+";"+type+";"+value+"\n")
                    file.close()
                    fenconfpro.destroy()
                    This.ButtonImportFiles['state'] = "active"
                    return
                
                props = {}
                
                fenconfpro = Toplevel(fenconf)
                fenconfpro.title("Configuration des propriétés")
                
                AddSTA = Button(fenconfpro, command=addSta, text="Ajout prop. statique", relief=RAISED, font=buttonfont)
                AddDYN = Button(fenconfpro, command=addDyn, text="Ajout prop. dynamique", relief=RAISED, font=buttonfont)
                
                AddSTA.grid(sticky="W",row=0, column=0,padx=3,pady=1)
                AddDYN.grid(sticky="E",row=0, column=1,padx=3,pady=1)
                
                Block = Canvas(fenconfpro, width=400, height=100)

                Block.grid(sticky="E",row=1, columnspan=2,padx=3,pady=1)
                
                if (os.path.exists(This.conf['dir'] + "Conf/properties.csv") == True):
                    fields = This.get_fieldsCSV()
                    file = open(This.conf['dir'] + "Conf/properties.csv", "r")
                    idx = 1
                    for prop in file.readlines():
                        line = prop.split("\n")[0].split(";")
                        if ( prop != "\n" ):
                            props[idx] = {}
                            props[idx]['type'] = line[2]
                            props[idx]['name'] = Entry(Block,bg="white",width=15)
                            props[idx]['name'].insert(0,line[0])
                            
                            props[idx]['formattix'] = Tix.StringVar()  
                            props[idx]['formatlist'] = Tix.ComboBox(Block, editable=1, dropdown=1, variable=props[idx]['formattix'], width=20, listwidth=15)
                            props[idx]['formatlist'].insert(0,"TXT")
                            props[idx]['formatlist'].insert(1,"NUM")
                            props[idx]['formatlist'].insert(2,"DATE")
                            props[idx]['formatlist'].entry.config(width=8, state='readonly')
                            props[idx]['formatlist'].slistbox.listbox.bind('<ButtonRelease-1>',partial(updateList,name=props[idx]['formatlist'],key=idx,word="format"))
                            props[idx]['format'] = line[1]
                            props[idx]['formattix'].set(line[1])
                                                       
                            props[idx]['name'].grid(sticky="W",row=idx, column=0,padx=3,pady=1)
                            props[idx]['formatlist'].grid(sticky="W",row=idx, column=1,padx=3,pady=1)
                            
                            
                            if ( props[idx]['type'] == "STA" ):
                                props[idx]['value'] = Entry(Block,bg="white",width=15)
                                props[idx]['value'].insert(0,unicode(line[3],"iso8859_1"))
                                props[idx]['value'].grid(sticky="W",row=idx, column=2,padx=3,pady=1)
                            else:
                                props[idx]['valuetix'] = Tix.StringVar()
                                props[idx]['valuelist'] = Tix.ComboBox(Block, editable=1, dropdown=1, variable=props[idx]['valuetix'], width=20, listwidth=50)
                                gen_listFields(props[idx]['valuelist'])
                                props[idx]['valuelist'].entry.config(width=15, state='readonly')
                                props[idx]['valuelist'].slistbox.listbox.bind('<ButtonRelease-1>',partial(updateList,name=props[idx]['valuelist'],key=idx,word="value"))
                                props[idx]['value'] = fields[int(line[3])]
                                props[idx]['valuetix'].set(fields[int(line[3])])
                                props[idx]['valuelist'].grid(sticky="W",row=idx, column=2,padx=3,pady=1)
                            idx=idx+1
                    file.close()
                
                Quit = Button(fenconfpro, text="Quitter", command=fenconfpro.destroy, relief=RAISED, font=buttonfont)
                Save = Button(fenconfpro, text="Sauver", command=saveProperties, relief=RAISED, font=buttonfont)

                Save.grid(sticky="W",row=20, column=0,padx=3,pady=1)
                Quit.grid(sticky="E",row=20, column=1,padx=3,pady=1)

                fenconfpro.withdraw()
                This.posfen(fenconfpro)
                return
            
            def loadPackage():
                fenconf.destroy()
                CommandOpenPackage(path=This.conf['dir'])
                return
            
            fenconf = Toplevel(fen)
       
            fenconf.title("Import Alfresco - Création d'un package")
            
            defaultfont = tkFont.Font(fenconf, size=10, family='Verdana', weight='bold')
            helpfont = tkFont.Font(fenconf, size=10, family='Verdana', slant='italic')
            buttonfont = tkFont.Font(fenconf, size=10, family='Verdana', weight='bold')
            
            PathCSV = Label(fenconf, fg=FGHIDDEN, width=60, font=defaultfont)
            Path = Label(fenconf, fg=FGHIDDEN, width=60, font=defaultfont)
            
            This.ButtonPkgPlace = Button(fenconf, text='Nouveau', command=PkgPlace, font=buttonfont, relief=GROOVE,width=30)
            This.ButtonImportCSV = Button(fenconf, text='Importer le CSV', command=ImportCSV, font=buttonfont, relief=GROOVE,width=30)
            This.ButtonImportCSV.config(state=DISABLED)
            This.ButtonImportFiles = Button(fenconf, text='Importer les documents', command=ImportFiles, font=buttonfont, relief=GROOVE,width=30)
            This.ButtonImportFiles.config(state=DISABLED)
            This.ButtonConfig = Button(fenconf, text='Gestion package.conf', command=Config, font=buttonfont, relief=GROOVE,width=30)
            This.ButtonConfig.config(state=DISABLED)
            This.ButtonAspects = Button(fenconf, text='Gestion aspects.conf', command=ConfigAspects, font=buttonfont, relief=GROOVE,width=30)
            This.ButtonAspects.config(state=DISABLED)
            This.ButtonProperties = Button(fenconf, text='Gestion properties.csv', command=ConfigProperties, font=buttonfont, relief=GROOVE,width=30)
            This.ButtonProperties.config(state=DISABLED)
            This.ButtonLoad = Button(fenconf, text='Exploiter', command=loadPackage, font=buttonfont, relief=GROOVE,width=30)
            This.ButtonLoad.config(state=DISABLED)
            
            if ( This.conf['open'] != "" ):
                This.ButtonPkgPlace['state'] = "disabled"
                This.ButtonConfig['state'] = "active"
                This.ButtonAspects['state'] = "active"
                This.ButtonProperties['state'] = "active"
                This.ButtonLoad['state'] = "active"
            
            This.ButtonPkgPlace.pack(side=TOP, anchor=W, expand=NO)
            This.ButtonImportCSV.pack(side=TOP, anchor=W, expand=NO)
            This.ButtonConfig.pack(side=TOP, anchor=W, expand=NO)
            This.ButtonAspects.pack(side=TOP, anchor=W, expand=NO)
            This.ButtonProperties.pack(side=TOP, anchor=W, expand=NO)
            This.ButtonImportFiles.pack(side=TOP, anchor=W, expand=NO)
            This.ButtonLoad.pack(side=TOP, anchor=W, expand=NO)
            
            fenconf.withdraw()
            This.posfen(fenconf)

            return
        
        def CommandClosePackage():
            This.conf['dir'] = ""
            This.conf['open'] = ""
            This.confpack = {}
            This.Logger.delete('1.0', END)
            This.ButtonGenerate['state'] = "disabled"
            This.ButtonTestHost['state'] = "disabled"
            This.ButtonUpload['state'] = "disabled"
            This.Force['state'] = "disabled"
            This.Bar['value'] = 0
            This.UpdateGuide("Etape 1 : Choisissez le dossier contenant le package")
            
        
        ### Fenêtre principale
        fen = Tix.Tk()

        BG = "#4F4F4F"
        FGHIDDEN = "#4F4F4F"
        FG = "#FFF"
        FGSTATE = "red"
        
        # Polices
        defaultfont = tkFont.Font(fen, size=10, family='Verdana', slant='italic')
        titlefont = tkFont.Font(fen, size=11, family='Verdana', weight='bold')
        menufont = tkFont.Font(fen, size=10, family='Verdana', weight='bold')
        submenufont = tkFont.Font(fen, size=10, family='Verdana')
        guidefont = tkFont.Font(fen, size=11, family='Verdana', weight='bold')
        buttonfont = tkFont.Font(fen, size=10, family='Verdana', weight='bold')
        forcefont = tkFont.Font(fen, size=10, family='Verdana', weight='bold')

        fen.title('Import Alfresco - Version '+This.version)
        fen.config(bg=BG, relief=GROOVE)

        FM = Frame(fen, bg="#eee")
        FM2 = Frame(fen, bg="#eee")
        FM3 = Frame(fen, bg="#eee")
        
        # Affichage des logs
        This.Logger = Text(fen, bg="#5f5f5f", fg="#ccc", width=110, height=40, highlightthickness="0")
        This.Logger.tag_configure('Error', foreground="#FF8888")
        This.Logger.tag_configure('Success', foreground="#88FF88")
        # Conteneur du traitement
        This.Treatment = Label(fen, bg=BG, fg=FGHIDDEN, width=60, font=defaultfont)
        
        menubar = Menu(fen)
        package = Menu(menubar, tearoff=0)
        package.add_command(label="Ouvrir", command=CommandOpenPackage, font=submenufont)
        package.add_command(label="Créer/Modifier", command=CommandCreatePackage, font=submenufont)
        package.add_command(label="Fermer", command=CommandClosePackage, font=submenufont)
        menubar.add_cascade(label="Package", menu=package, font=menufont)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Configuration", command=CommandConfGlobale, font=submenufont)
        filemenu.add_command(label="Quitter", command=fen.destroy, font=submenufont)
        menubar.add_cascade(label="Général", menu=filemenu, font=menufont)


        # Conteneur du package
        PathDir = Label(fen, bg=BG, fg=FGHIDDEN, width=60, font=defaultfont)

        # Traitement des documents
        # Bouton du traitement
        This.ButtonGenerate = Button(FM, text='Générer le package', command=CommandGenerate, font=buttonfont, relief=GROOVE)
        This.ButtonGenerate.config(state=DISABLED)

        # Import dans Alfresco
        # Bouton d'upload
        This.ButtonUpload = Button(FM, text='Importer (mode Bulk Import)', command=CommandUpload, font=buttonfont, relief=GROOVE)
        This.ButtonUpload.config(state=DISABLED)
        
        This.ButtonTestHost = Button(FM, text='Test connexions', command=CommandTestHost, font=buttonfont, relief=GROOVE)
        This.ButtonTestHost.config(state=DISABLED)
        
        This.var1 = IntVar()
        This.Force = Checkbutton(FM2, text = "Mise à jour uniquement (CMIS)", highlightthickness="0", font="forcefont", variable = This.var1 , command=ChangeMode)
        This.Force.config(state=DISABLED)

        # Affichage du guide
        This.Guide = Text(FM3, bg="ivory", fg="#222", width=110, height=1.4, font=guidefont)

        # Bar de progression
        s = ttk.Style()
        s.theme_use('clam')
        s.configure("red.Horizontal.TProgressbar", foreground='red', background='red')
        This.Bar = ttk.Progressbar(FM3, style="red.Horizontal.TProgressbar", orient="horizontal", length=300, mode="determinate", maximum=300)

        # Bouton de nettoyage de l'écran des logs
        #"Clear = Button(FM2, text="Nettoyer les logs", command=CommandClearLog, font=buttonfont, relief=RAISED)

        # Placements
        This.ButtonGenerate.pack(side=LEFT, anchor=W, fill=X, expand=YES)
        This.ButtonTestHost.pack(side=LEFT, anchor=W, fill=X, expand=YES)
        This.ButtonUpload.pack(side=LEFT, anchor=W, fill=X, expand=YES)
        FM.pack(fill=X)
        FM2.pack(fill=X)
        FM3.pack(fill=X)
        This.Logger.pack(side=BOTTOM, fill=X)
        This.Guide.pack(side=BOTTOM, fill=X)
        This.Bar.pack(side=BOTTOM, fill=X)
        #This.Mode.pack(side=RIGHT, fill=X)
        This.Force.pack(side=RIGHT, fill=X)
        #Clear.pack(side=RIGHT, fill=X)

        #Clear.pack(side=BOTTOM, anchor=W, fill=X, expand=YES)

        This.Guide.tag_configure('tag-center', justify='center')
        This.UpdateGuide("Etape 1 : Choisissez le dossier contenant le package")

        #Center(fen)
        fen.config(menu=menubar)
        
        # Taille et position de la fenêtre
        fen.resizable(width=False, height=False)
        
        This.posfen(fen, 1024, 660)
        
        fen.mainloop()

        return

        
    def testWkspace(This, path):
        try:
            client = CmisClient(This.conf['url'], This.conf['user'], This.conf['password'])
            repo = client.defaultRepository

            Folder = repo.getObjectByPath(path)
            return [True,path + "' : OK", "Success"]
        except Exception, e:
            return [False,path + "' : Introuvable", ""]

    def testCMIS(This):
        try:
            client = CmisClient(This.conf['url'], This.conf['user'], This.conf['password'])

            repo = client.defaultRepository
            REQ = "select * from cmis:folder where cmis:name = 'Sites'"

            results = repo.query(REQ)
            
            if (len(results) != 0):
                return str(results[0].id)
            else:
                return False
        except Exception, e:
            This.logger(str(e),"Error")
            return False
        
    def testBulkImport(This):
        try:
            rh = RESTHelper()
            rh.login(This.conf['user'], This.conf['password'], This.conf['host'], 8080)
            data = rh.statusbulkimport()
            return True
        except Exception, e:
            This.logger(str(e),"Error")
            return False
    
    def getStatusBulkImport(This):
        try:
            rh = RESTHelper()
            rh.login(This.conf['user'], This.conf['password'], This.conf['host'], 8080)
            data = rh.statusbulkimport()
            return json.load(data)
        except Exception, e:
            This.logger(str(e),"Error")
            return False
    
    def OpenHost(This, mode, silence):
        This.Logger.delete('1.0', END)
        if ( This.conf['host'] != "" ):
            This.conf['url'] = This.conf['urltemp'].replace("__HOST__", This.conf['host'])
            
            testbulkimport = This.testBulkImport()

            if ( testbulkimport ):
                This.logger("Test connexion Bulk Import Tool : OK","Success",silence)
            else:
                This.logger("Test connexion Bulk Import Tool : Error","Error",silence)
                return False
            
            if (os.path.exists(This.conf['dir'] + "Conf/package.conf")):
                testhost = This.testCMIS()
                
                if (testhost != False ):
                    This.logger("Test connexion CMIS Alfresco ("+This.conf['host']+") : OK","Success",silence)
                    This.logger("\nTest des destinations du CSV :\n","",silence)

                    pathlist = {}
                    WKTEST = True
                    for fid in This.fileinfo:

                        path = This.fileinfo[fid]['upath']

                        if ( path not in pathlist ):
                            pathlist[path] = True
                            test = This.testWkspace(path)
                            This.logger(test[1], test[2],silence)
                            if ( test[0] == False ):
                                WKTEST = False
                else:
                    WKTEST = False

                if ( WKTEST == True ):
                    if ( mode == "BULKIMPORTTOOL" ):
                        This.ButtonUpload['state'] = "active"
                        This.Force['state'] = "active"
                    else:
                        This.ButtonUpload['state'] = "active"
                        This.var1.set(1)
                        This.Force['state'] = "active"
                    return True
                else:
                    This.Force['state'] = "disabled"
                    This.var1.set(0)
                    This.logger("\nTest des destinations : dossiers manquants (sans conséquences en mode Bulk Import Tool)","",silence)
                    This.logger("\nPour le mode CMIS (mise à jour), le mode Bulk Import Tool doit être lancé en premier)","",silence)
                    This.ButtonUpload['text'] = "Importer (mode Bulk Import)"
                    if ( mode == "BULKIMPORTTOOL" ):
                        return True
                    else:
                        return False
        else:
            This.logger("\nLe serveur hôte Alfresco n'est pas configuré","Error")
            

    # Get files informations from CSV
    def parse_csv(This):
        This.fileinfo = {}

        csvfile = open(This.conf['dir'] + 'Conf/list.csv', "rb")
        reader = csv.reader(csvfile, delimiter=';', quotechar='"')
        firstline = 1    
        for row in reader:
            if (firstline == 1):
                firstline = 0
            else:
                fid = str(row[0])
                name = unicode(row[This.confpack['name']], "iso8859_1")
                newname = unicode(row[This.confpack['newname']], "iso8859_1")
                bname = row[This.confpack['name']]
                bnewname = row[This.confpack['newname']]
                title = unicode(row[This.confpack['title']], "iso8859_1")
                description = unicode(row[This.confpack['desc']], "iso8859_1")
                tags = unicode(row[This.confpack['tags']], "iso8859_1").split(",")
                path = unicode(row[This.confpack['path']], "iso8859_1")
                upath = row[This.confpack['path']].decode('iso8859_1').encode('utf-8')

                idx = "F" + fid

                This.fileinfo[idx] = {}
                This.fileinfo[idx]['id'] = fid
                This.fileinfo[idx]['title'] = title
                This.fileinfo[idx]['description'] = description
                This.fileinfo[idx]['tags'] = tags
                This.fileinfo[idx]['name'] = name
                This.fileinfo[idx]['newname'] = newname

                # Version sans unicode
                This.fileinfo[idx]['bname'] = bname
                This.fileinfo[idx]['bnewname'] = bnewname

                This.fileinfo[idx]['path'] = path
                This.fileinfo[idx]['upath'] = upath

                This.fileinfo[idx]['properties'] = {}

                This.fileinfo[idx]['properties'] = This.get_properties(This.fileinfo[idx]['properties'], row)

        csvfile.close()

        return This.fileinfo

    def get_aspects(This):
        file = open(This.conf['dir'] + "Conf/aspects.conf", "r")
        Aspects = []
        for aspect in file.readlines():
            if ( aspect != "\n"):
                Aspects.append(aspect.rstrip())
        file.close()
        return Aspects

    def get_properties(This,tab, data):
        csvfile = open(This.conf['dir'] + 'Conf/properties.csv', "rb")
        reader = csv.reader(csvfile, delimiter=';', quotechar='"')

        for row in reader:
            if (row[2] == "DYN"):
                tab[row[0]] = data[int(row[3])]
            else:
                tab[row[0]] = row[3]
                
            if (row[1] == "DATE"):
                sdate = tab[row[0]].split("/")
                tab[row[0]] = sdate[2] + "-" + sdate[1] + "-" + sdate[0] + "T00:00:00.000+00:00"
            if (row[1] == "TXT"):
                tab[row[0]] = unicode(tab[row[0]], "iso8859_1")
            if (row[1] == "NUM"):
                tab[row[0]] = int(tab[row[0]])

        csvfile.close()
        return tab

    def generatepack(This):  
        counter = 1
        This.Bar['maximum'] = len(This.fileinfo)
        val = 1
        pval = 0
        
            
        for fid in This.fileinfo:
            name = This.fileinfo[fid]['name']
            newname = This.fileinfo[fid]['newname']
            title = This.fileinfo[fid]['title']
            description = This.fileinfo[fid]['description']
            path = This.fileinfo[fid]['path']
            
            if ( os.path.exists(This.conf['dir'] + This.confpack['PKGID'] + path +'/'+newname) ):
                This.logger(str(counter) + " - " + newname + " existe déjà","")
                
                pval = pval + val
                This.Bar['value'] = pval

                This.logger(str(counter) + " - " + newname,"")
                counter = counter + 1
            else :
                pdf = PdfFileReader(open(This.conf['dir'] + "Orig/" + name, 'rb'))

                if (pdf.getIsEncrypted()):
                    pdf.decrypt('')

                data = {u'/Title':u'%s' % title, u'/Subject':u'%s' % description}

                This.add_metadata(name, path, newname, data)
                LOG = u"Traitement des métadonnées du document '%s'" % (newname)

                This.genXMLProperties(fid)

                pval = pval + val
                This.Bar['value'] = pval

                This.logger(str(counter) + " - " + newname,"")
                counter = counter + 1

        pval = pval + val        
        This.Bar['value'] = pval

        if ( This.Bar['value'] >= This.Bar['maximum'] ):
            return True
        else:
            return False
        
    def genXMLProperties(This,fid):
        id = This.fileinfo[fid]['id']

        xmlfile = open(This.conf['dir'] + This.confpack['PKGID'] +This.fileinfo[fid]['path']+'/'+This.fileinfo[fid]['newname']+'.metadata.properties.xml', "wb")
        
        xmlfile.write('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE properties SYSTEM "http://java.sun.com/dtd/properties.dtd">\n<properties>\n')
        xmlfile.write('<entry key="separator"> # </entry>\n')
        xmlfile.write('<entry key="type">cm:content</entry>\n')
        
        Aspects = This.get_aspects()
        
        Aspects.append("ialf:package")
        
        count = len(Aspects)
        c=1
        listasp = ""
        
        for aspect in Aspects:
            listasp = listasp + aspect
            if ( c < count ):
                listasp = listasp + " # "
            c=c+1 
        
        xmlfile.write('<entry key="aspects">%s</entry>\n' % listasp)
        
        for prop in This.fileinfo[fid]['properties']:
            propname = prop
            propvalue = This.fileinfo[fid]['properties'][prop]

            xmlfile.write(u'<entry key="%s">%s</entry>\n' % (propname, propvalue))
            
        xmlfile.write(u'<entry key="ialf:packageid">%s</entry>\n' % (This.confpack['PKGID']+"_"+id))
        
        xmlfile.write('</properties>\n')
        
        xmlfile.close()
        
        return

    def add_metadata(This,name, path, newname, data):
        merger = PdfFileMerger()

        with open(This.conf['dir'] + 'Orig/%s' % name, 'rb') as f0:
            merger.append(f0)

        merger.addMetadata(data)

        if not os.path.exists(os.path.dirname(This.conf['dir'] + This.confpack['PKGID'] + '%s/%s' % (path,newname))):
            os.makedirs(os.path.dirname(This.conf['dir'] + This.confpack['PKGID'] + '%s/%s' % (path,newname)))

        with open(This.conf['dir'] + This.confpack['PKGID'] + '%s/%s' % (path,newname), 'wb') as f1:
            merger.write(f1)

    def upload(This, mode):  
        counter = 1
        This.Bar['maximum'] = len(This.fileinfo)
        This.Bar['value'] = 0 
        This.logger("\nDébut de l'import vers "+This.conf['host'],"")
        FileMissing = False
        for fid in This.fileinfo:
            newname = This.fileinfo[fid]['newname']
            #This.cmisCreate(newname, This.fileinfo[fid], force, counter)
            RETURN = This.updateCmis(This.fileinfo[fid], counter)
            if ( RETURN == False ):
                This.logger("Document '"+newname+"' manquant","Error")
                FileMissing = True
            counter=counter+1

        if ( This.Bar['value'] >= This.Bar['maximum'] ):
            if ( FileMissing ):
                This.logger("\nATTENTION: des documents du package sont manquant dans Alfresco. Vous devez sans doute relancer l'import en mode Bulk Import Tool.","Error")
                return "Missing"
            return True
        else:
            return False

    def getSitesId(This):
        try:
            client = CmisClient(This.conf['url'], This.conf['user'], This.conf['password'])

            repo = client.defaultRepository
            REQ = "select * from cmis:folder where cmis:name = 'Sites'"

            results = repo.query(REQ)

            if (len(results) != 0):   
                return str(results[0].id)
            else:
                return False
        except Exception, e:
            This.logger(str(e),"Error")
            return False

    def cmisCreate(This, name, onefileinfo, force, counter):
            id = onefileinfo['id']
            pathdst = onefileinfo['path']

            client = CmisClient(This.conf['url'], This.conf['user'], This.conf['password'])

            repo = client.defaultRepository

            Folder = repo.getObject("workspace://SpaceStore/" + This.conf['wkimports'])

            REQ = "select * from ialf:package where ialf:packageid = '" + This.confpack['PKGID'] + "_" + id + "'"

            results = repo.query(REQ)

             # Création du document
            if (len(results) == 0):
                try:
                    LOG = u"Création du document %s" % name
                    File = open(This.conf['dir'] + This.confpack['PKGID'] + pathdst + "/" + name, 'r')
                    Doc = Folder.createDocument(This.confpack['PKGID'] + "_" + id+".pdf", contentFile=File)
                    File.close()
                    This.logger(LOG,"")
                    REQ = "select * from cmis:document where cmis:name = '" + This.confpack['PKGID'] + "_" + id+".pdf" + "'"
                    force = 1
                except Exception, e:
                    This.logger(u"Problème de création du document %s" % name,"Error")
                    This.logger(str(e),"Error")
            else:
                LOG = u"%s déjà existant" % name
                This.logger(LOG,"")

            if (force == 1):
               This.updateCmis(onefileinfo, counter)

            This.Bar['value'] = counter
            
    def updateCmis(This, onefileinfo, counter):
        id = onefileinfo['id']

        client = CmisClient(This.conf['url'], This.conf['user'], This.conf['password'])

        repo = client.defaultRepository

        REQ = "select * from ialf:package where ialf:packageid = '" + This.confpack['PKGID'] + "_" + id + "'"

        results = repo.query(REQ)
        
        if (len(results) == 0):
            return False
        
        # Aspects
        if (len(results) != 0):
            try:
                LOG = u"--> Application des aspects du document %s" % onefileinfo['newname']
                objectId = str(results[0].id)
                Doc = repo.getObject("workspace://SpaceStore/" + objectId)
                Aspects = This.get_aspects()
                for aspect in Aspects:
                    Doc.addAspect("P:" + aspect)
                Doc.addAspect("P:ialf:package")
                This.logger(LOG,"")
            except Exception, e:
                This.logger(u"Problème d'ajout des aspects pour %s" % onefileinfo['newname'],"Error")
                This.logger(str(e),"Error")
                return False

        # Propriétés
        results = repo.query(REQ)
        if (len(results) != 0):
            try:
                LOG =  u"--> Création des propriétés du document %s" % onefileinfo['newname']
                objectId = str(results[0].id)
                Doc = repo.getObject("workspace://SpaceStore/" + objectId)
                props = {}

                This.logger(LOG,"")

                for prop in onefileinfo['properties']:
                    propname = prop
                    propvalue = onefileinfo['properties'][prop]

                    props[propname] = propvalue

                props["ialf:packageid"] = This.confpack['PKGID']+"_"+id
                props["cmis:name"] = onefileinfo['newname']
                Doc.updateProperties(props)
            except Exception, e:
                This.logger(u"Problème d'ajout des propriétés pour %s" % onefileinfo['newname'],"Error")
                This.logger(str(e),"Error")
                return False

        if (len(results) != 0):
            objectId = str(results[0].id)
            try:
                for tag in onefileinfo['tags']:
                    if (tag != ""):
                        split_objId = objectId.split(";")
                        This.add_tags(split_objId[0], tag)
                This.logger(u"--> Ajout des tags pour %s" % onefileinfo['newname'],"")
            except Exception, e:
                This.logger(u"Problème d'ajout des tags pour %s" % onefileinfo['newname'],"Error")
                This.logger(str(e),"Error")
                return False
        
        # Déplacement du document
#        try:
#            REQ = "select * from ialf:package where ialf:packageid = '" + This.confpack['PKGID'] + "_" + id + "'"
#            results = repo.query(REQ)
#            if (len(results) != 0):
#                objectId = str(results[0].id)
#                Doc = repo.getObject("workspace://SpaceStore/" + objectId)
#                parent = Doc.getObjectParents().getResults()[0]
#                if (str(parent) == This.conf['wkimports']):
#                    src = repo.getObject("workspace://SpaceStore/" + This.conf['wkimports'])
#                    dst = repo.getObject("workspace://SpaceStore/" + wkspacedst)
#                    Doc.move(src, dst)
#        except Exception, e:
#                This.logger(u"Problème de déplacement de %s" % name,"Error")
#                This.logger(str(e),"Error")
#                return False

        This.Bar['value'] = counter
            
    def transfertPack(This):
        import scpclient
        try:
            client = paramiko.SSHClient()
            client.load_system_host_keys()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(This.conf['host'], 22, This.conf['user_ssh'], This.conf['password_ssh'])
            
            createdir = False
            try:
                with closing(scpclient.Read(client.get_transport(), This.conf['importdir']+"/"+This.confpack['PKGID'])) as scp:
                    scp.receive('Sites.metadata.properties.xml')
                    distdir = This.conf['importdir']+"/"+This.confpack['PKGID']
                    if ( distdir != "/"):
                        client.exec_command("rm -fr "+distdir, timeout=60)
                        createdir=True
                    else:
                        This.logger("Suppression du package distant impossible","Error")
                        return False
            except:
                createdir = True
            
            if ( createdir ):
                dirempty=This.conf['dir']+"EMPTY"
                if (os.path.exists(dirempty) == False):
                    os.mkdir(dirempty)
                with closing(scpclient.WriteDir(client.get_transport(), This.conf['importdir']+"/"+This.confpack['PKGID'])) as scp:
                    scp.send_dir(dirempty, override_mode=True, preserve_times=True)
                os.rmdir(dirempty)
                
                with closing(scpclient.Write(client.get_transport(), This.conf['importdir']+"/"+This.confpack['PKGID'])) as scp:
                    scp.send_file(This.conf['dir'] + This.confpack['PKGID'] + "/Sites.metadata.properties.xml", remote_filename="Sites.metadata.properties.xml")
            
            pathdir = This.conf['dir'] + This.confpack['PKGID'] + "/Sites"
            with closing(scpclient.WriteDir(client.get_transport(), This.conf['importdir']+"/"+This.confpack['PKGID'])) as scp:
                scp.send_dir(pathdir, override_mode=True, preserve_times=True)
                
            client.close()
            return True
        except Exception, e:
            This.logger(str(e),"Error")
            return False
        
    def returnStatus(This, result):
        tag=""
        
        This.logger("\n - Package                       : "+str(result['sourceParameters']['Source Directory']),"")
        
        status = "En attente"
            
        if ( str(result['processingState']) == "Failed" ):
            status = "Echoué"
            tag="Error"

        if ( str(result['processingState']) == "Succeeded" ):
            status = "Succès"
            tag="Success"
        
        This.logger(" - Statut                        : "+status,tag)
        This.logger(" - Durée du traitement           : "+str(result['duration']),"")
        This.logger(" - Taille doc. importés (octets) : "+str(result['targetCounters']['Bytes imported']['Count']),"")
        
        if ( result['inProgress'] == False ):
            progress = "Terminé"
        else:
            progress = "En cours"
        
        This.logger(" - Progression                   : "+progress+"\n",tag)
            
        return result['inProgress']
    
    def initiateBulkImport(This):
        This.logger("\nTransfert du package sur "+This.conf['host']+" (Bulk Import Tool). Patientez...\n","")
        if ( This.transfertPack() ):
            This.logger("Transfert du package sur "+This.conf['host']+" OK\n","Success")
            
            rh = RESTHelper()
            rh.login(This.conf['user'], This.conf['password'], This.conf['host'], 8080)

            try:
                rh.initiateBulkImport(This.conf['importdir']+"/"+This.confpack['PKGID']+"/","/")
                This.logger("Initialisation de l'import (Bulk Import Tool)","Success")
                
                result = This.getStatusBulkImport()
                while ( This.returnStatus(result) ):
                    time.sleep(2)
                    This.Logger.delete('8.0', END)
                    result = This.getStatusBulkImport()
             
                return True
            except Exception, e:
                This.logger(str(e),"Error")
                return False
        else:
            This.logger("Transfert du package sur "+This.conf['host']+" Error\n","Error")

    def add_tags(This, objectId, tag):
        rh = RESTHelper()
        rh.login(This.conf['user'], This.conf['password'], This.conf['host'], 8080)

        rh.addTag("workspace", "SpacesStore", objectId, tag)

        rh.logout
        return


