#!/usr/bin/python
# coding: utf-8
# By Sébastien Brière - 2016
# Importations de documents dans Alfresco
# Sous licence GLP

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
import xlrd
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
    def get_conf(self):
        config = ConfigParser.ConfigParser()
        
        file = open(self.dir_path+"/importalf.conf", "rb")

        config.readfp(file)
        
        file.close()
        
        conf = {}
        LOG = "Paramètre introuvable : "
        try:
            conf['urltemp'] = config.get("GLOBAL", "url")
            if ( conf['urltemp'] == "" ):
                self.logger(LOG+"urltemp","Error")
        except:
            self.logger(LOG+"urltemp","Error")
            conf["urltemp"] = ""
        
        try:
            conf['user'] = config.get("GLOBAL", "user")
            if ( conf['user'] == "" ):
                self.logger(LOG+"user","Error")
        except:
            self.logger(LOG+"user","Error")
            conf["user"] = ""
        
        try:
            conf['password'] = config.get("GLOBAL", "password")
            if ( conf['password'] == "" ):
                self.logger(LOG+"password","Error")
        except:
            self.logger(LOG+"password","Error")
            conf["password"]
        
        try:
            conf['host'] = config.get("GLOBAL", "host")
            if ( conf['host'] == "" ):
                self.logger(LOG+"host","Error")
        except:
            self.logger(LOG+"host","Error")
            conf["host"] = ""
        
        try:
            conf['importdir'] = config.get("GLOBAL", "bulkimportdir")
            if ( conf['importdir'] == "" ):
                self.logger(LOG+"importdir","Error")
        except:
            self.logger(LOG+"importdir","Error")
            conf["importdir"] = ""
            
        try:
            conf['user_ssh'] = config.get("GLOBAL", "user_ssh")
            if ( conf['user_ssh'] == "" ):
                self.logger(LOG+"user_ssh","Error")
        except:
            self.logger(LOG+"user_ssh","Error")
            conf["user_ssh"] = ""
        
        try:
            conf['password_ssh'] = config.get("GLOBAL", "password_ssh")
            if ( conf['password_ssh'] == "" ):
                self.logger(LOG+"password_ssh","Error")
        except:
            self.logger(LOG+"password_ssh","Error")
            conf["password_ssh"] = "" 
            
        conf['port'] = "8080"    
        return conf

    # Get Package Config
    def get_confpack(self):
        config = ConfigParser.ConfigParser()
        file = open(self.conf['dir'] + "Conf/package.conf", "rb")
        config.readfp(file)
        listformat = ""
        if ( "listformat" in self.confpack ):
            listformat = self.confpack['listformat']
        listpath = ""
        if ( "listpath" in self.confpack ):
            listpath = self.confpack['listpath']

        self.confpack = {}
        self.confpack['listformat'] = listformat
        self.confpack['listpath'] = listpath
        self.confpack['name'] = int(config.get("PACKAGE", "OLDNAME"))
        self.confpack['newname'] = int(config.get("PACKAGE", "NEWNAME"))
        self.confpack['title'] = int(config.get("PACKAGE", "TITLE"))
        self.confpack['desc'] = int(config.get("PACKAGE", "DESC"))
        self.confpack['tags'] = int(config.get("PACKAGE", "TAGS"))
        self.confpack['path'] = int(config.get("PACKAGE", "DESTPATH"))
        return self.confpack

    # Logger
    def logger(self, txt, tag, silence=False):
        if ( silence == False ):
            self.Logger.insert(END," █ " + txt + "\n",tag)
            self.Logger.see("end")
            self.Treatment.update()

    # Generate and get the packageId
    def get_packageid(self):
        if (os.path.exists(self.conf['dir'] + "Conf/packageId")):
            file = open(self.conf['dir'] + "Conf/packageId", "rb")
            PKGID = file.read()
            file.close()
        else:
            PKGID = str(uuid.uuid4())
            file = open(self.conf['dir'] + "Conf/packageId", "wb")
            file.write(PKGID)
            file.close()

        return PKGID

    def UpdateGuide(self,txt):
            self.Guide.delete('1.0', END)
            self.Guide.insert("1.0", txt ,"tag-center")
    
    def askYesNo(self, title, message, parent):
        return tkMessageBox.askquestion(title,message,parent=parent)
    
    def messageShow(self, title, message, parent):
        return tkMessageBox.showwarning(title,message,parent=parent)
    
    def getIndex(self,value,dicto):
        return dicto.keys()[dicto.values().index(value)]
     
    def gen_dialog(self):
        def CommandClearLog():
            self.Logger.delete('1.0', END)
        
        def CommandGenerate():
            self.Logger.delete('1.0', END)
  
            if (os.path.exists(self.conf['dir'] + "Conf/package.conf")):
                if ( os.path.exists(self.conf['dir']+"/"+self.confpack['PKGID']) == False ):
                    os.mkdir(self.conf['dir']+"/"+self.confpack['PKGID'])
                else:
                    shutil.rmtree(self.conf['dir']+"/"+self.confpack['PKGID'])
                    os.mkdir(self.conf['dir']+"/"+self.confpack['PKGID'])
                    
#                initfile=open(self.conf['dir'] + self.confpack['PKGID'] +"/Sites.metadata.properties.xml","wb")
#                initfile.close()
                RESULT = self.generatepack()
                if ( RESULT == True ):
                    self.ButtonGenerate['state'] = "disabled"
                    self.ButtonUpload['state'] = "active"
                    self.Force['state'] = "active"
                    self.logger("Génération des documents réussie","Success")
                    self.UpdateGuide("Etape 3 : Importez les documents")    

        def CommandOpenPackage(path="",silence=0):
            self.Logger.delete('1.0', END)
            
            try:
                if ( path == "" ):
                    # Dialog Box for choose dir package
                    PackageDir = tkFileDialog.askdirectory(initialdir=self.dir_path, title="Choisir le répertoire")
                    PathDir['text'] = PackageDir
                    if ( PathDir['text'] == "" ):
                        return
                    self.conf['dir'] = PathDir['text'] + "/"
                else:
                    self.conf['dir'] = path

                check_package_conf = False
                check_aspects_conf = False
                check_properties_conf = False

                if (os.path.exists(self.conf['dir'] + "Conf/package.conf")):
                    self.confpack = self.get_confpack()
                else:
                    self.logger("","",silence)
                    self.logger("Le fichier package.conf est introuvable","Error",silence)

                if (os.path.exists(self.conf['dir'] + "Conf/list.csv") or os.path.exists(self.conf['dir'] + "Conf/list.xls") or os.path.exists(self.conf['dir'] + "Conf/list.xlsx")):
                    self.fileinfo = self.parse_csv()
                    self.logger("Nombre de documents : "+str(len(self.fileinfo)),"",silence)

                    fields = self.get_fieldsCSV()

                    name = fields[self.confpack['name']]
                    newname = fields[self.confpack['newname']]
                    title = fields[self.confpack['title']]
                    description = fields[self.confpack['desc']]
                    tags = fields[self.confpack['tags']]
                    path = fields[self.confpack['path']]

                    self.logger("","",silence)
                    self.logger("Vérification correspondance des champs :","",silence)
                    self.logger("","",silence)
                    self.logger("  Attributs   | Champs CSV","",silence)
                    self.logger("  ____________|___________    ","",silence)
                    self.logger("  OLDNAME     | "+name,"",silence)
                    self.logger("  NEWNAME     | "+newname,"",silence)
                    self.logger("  TITLE       | "+title,"",silence)
                    self.logger("  DESC        | "+description,"",silence)
                    self.logger("  TAGS        | "+tags,"",silence)
                    self.logger("  DESTPATH    | "+path,"",silence)
                    
                    check_package_conf = True
                                  
                    for line in self.fileinfo:
                        if ( os.path.exists(self.conf['dir'] + "Orig/" +self.fileinfo[line]['name']) == False):
                            self.logger("","",silence)
                            self.logger("Document du CSV manquant dans le répertoire","Error",silence)
                            check_package_conf = False
                            break
                else:
                    self.logger("","",silence)
                    self.logger("Le fichier list.csv est introuvable","Error",silence)
                    self.logger("","",silence)
                    
                if (os.path.exists(self.conf['dir'] + "Conf/aspects.conf") and check_package_conf ):    
                    self.logger("","",silence)
                    self.logger("Liste des aspects :","",silence)
                    self.logger("","",silence)

                    file = open(self.conf['dir'] + "Conf/aspects.conf", "r")
                    for aspect in file.readlines():
                        self.logger("--> "+aspect.rstrip(),"",silence)
                    file.close()
                    check_aspects_conf = True
                else:
                    if ( check_package_conf ):
                        self.logger("","",silence)
                        self.logger("Le fichier aspects.conf est introuvable","Error",silence)
                        self.logger("","",silence)

                if (os.path.exists(self.conf['dir'] + "Conf/properties.csv") and check_package_conf ):
                    self.logger("","",silence)
                    self.logger("Liste des propriétés :","",silence)
                    self.logger("","",silence)

                    csvfile = open(self.conf['dir'] + 'Conf/properties.csv', "rb")
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

                        self.logger("--> : "+name+valuedetail,"",silence)
                        check_properties_conf = True
                    csvfile.close()
                else:
                    if ( check_package_conf ):
                        self.logger("","",silence)
                        self.logger("Le fichier properties.conf est introuvable","Error",silence)
                        self.logger("","",silence)

                if ( check_package_conf and check_aspects_conf and check_properties_conf ):
                    PKGID = self.get_packageid()

                    self.confpack['PKGID'] = PKGID
                    self.logger("","",silence)
                    self.logger("Package ID : "+self.confpack['PKGID'], "Success",silence)
                    self.logger(self.conf['dir'] + " : OK.", "Success",silence)

                    self.ButtonGenerate['state'] = "active"
                    self.ButtonTestHost['state'] = "active"
                    self.UpdateGuide("Etape 2 : Générez le package")
                    self.conf['open'] = self.confpack['PKGID']
                    return True
                else:
                    self.logger("","",silence)
                    self.logger(self.conf['dir'] + " : Erreur.", "Error",silence)
                    self.ButtonGenerate['state'] = "disabled"
                    self.ButtonTestHost['state'] = "disabled"
                    self.ButtonUpload['state'] = "disabled"
                    return False
            except Exception, e:
                print str(e)
                return False
        
        def CommandUpload():
            mode = "BULKIMPORTTOOL"
            
            if ( self.var1.get() == 1 ):
                mode = "CMIS"
            
            TestHost = self.OpenHost(mode,True)
            
            if ( TestHost ):
                if (os.path.exists(self.conf['dir'] + "Conf/package.conf")):
                    if ( mode == "CMIS" ):
                        RESULT = self.upload(mode)
                    else:
                        RESULT = self.initiateBulkImport()
                    if ( RESULT == True ):
                        self.logger("Import des documents réussie","Success")
                        self.UpdateGuide("Terminé")
                        self.ButtonUpload['state'] = "active"
                        self.Force['state'] = "active"
                    else:
                        if ( RESULT == "Missing" ):
                            self.logger("Import des documents réussi, mais quelques documents manquants","Success")
                        else:
                            self.logger("Import des documents échoué","Error")
                    
        def ChangeMode():
            if ( self.var1.get() == 0 ):
                self.ButtonUpload['text'] = "Importer (mode Bulk Import)"
            else:
                self.ButtonUpload['text'] = "Mettre à jour (mode CMIS)"
        
        def CommandTestHost():
            mode = "BULKIMPORTTOOL"
            TestHost = self.OpenHost(mode, False)
        
        def CommandConfGlobale():
            def SaveConf():
                file = open(self.dir_path+"/importalf.conf", "wb")
                file.write("[GLOBAL]\n")
                file.write("url="+UrlTemp.get()+"\n")
                file.write("user="+User.get()+"\n")
                file.write("password="+Cred.get()+"\n")
                file.write("host="+Host.get()+"\n")
                file.write("bulkimportdir="+BulkImportDir.get()+"\n")
                file.write("user_ssh="+UserSSH.get()+"\n")
                file.write("password_ssh="+CredSSH.get()+"\n")
                file.close()
                self.conf = self.get_conf()
                self.conf['dir'] = PathDir['text'] + "/"
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
            UrlTemp.insert(END,self.conf['urltemp'])
            
            User = Entry(fenconf,bg="white",width=15)
            User.insert(END,self.conf['user'])
            
            Cred = Entry(fenconf,show="*",bg="white",width=15)
            Cred.insert(END,self.conf['password'])
            
            UserSSH = Entry(fenconf,bg="white",width=15)
            UserSSH.insert(END,self.conf['user_ssh'])
            
            CredSSH = Entry(fenconf,show="*",bg="white",width=15)
            CredSSH.insert(END,self.conf['password_ssh'])
            
            Host = Entry(fenconf,bg="white",width=45)
            Host.insert(END,self.conf['host'])
            
            BulkImportDir = Entry(fenconf,bg="white",width=45)
            BulkImportDir.insert(END,self.conf['importdir'])
            
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
            
            fenconf.withdraw()
            l = self.posfen(fenconf,bottom=True)
            self.displayBottom(fenconf,l,10,partial(SaveConf))
            return
        
        def gen_listFields(box):
            fields = self.get_fieldsCSV()
            for field in fields:
                box.insert(field,fields[field])
            return
        
        def CommandCreatePackage():
            def PkgPlace():
                Place = tkFileDialog.askdirectory(parent=fenconf,title="Choisir le CSV", initialdir=self.dir_path) 
                if ( Place == () ):
                    return
                try:
                    self.conf['dir'] = Place + "/"
                    
                    if (os.path.exists(self.conf['dir'] + "Conf") == False):
                        os.mkdir(self.conf['dir']+"Conf")
                    if (os.path.exists(self.conf['dir'] + "Orig") == False):
                        os.mkdir(self.conf['dir']+"Orig")
                    self.get_packageid()
                    
                    
                    self.ButtonPkgPlace['state'] = "disabled"
                    if ( CommandOpenPackage(path=self.conf['dir'],silence=1) ):
                        self.ButtonImportCSV['state'] = "active"
                        self.ButtonConfig['state'] = "active"
                        self.ButtonAspects['state'] = "active"
                        self.ButtonProperties['state'] = "active"
                        self.ButtonImportFiles['state'] = "active"
                        self.ButtonLoad['state'] = "active"
                    else:
                        self.ButtonImportCSV['state'] = "active"
                    return True
                except Exception, e:
                    self.logger(str(e),"Error")
                    return False
            
            def ImportCSV():
                try:
                    if (os.path.exists(self.conf['dir'] + "Conf/list.csv") == True or os.path.exists(self.conf['dir'] + "Conf/list.xls") == True or os.path.exists(self.conf['dir'] + "Conf/list.xlsx") == True):
                        if ( os.path.exists(self.conf['dir'] + "Conf/list.csv") ):
                            self.confpack['listformat'] = "CSV"
                            self.confpack['listpath'] = self.conf['dir'] + "Conf/list.csv"
                        else:
                            self.confpack['listformat'] = "XLS"
                            if ( os.path.exists(self.conf['dir'] + "Conf/list.xls") == True ):
                                self.confpack['listpath'] = self.conf['dir'] + "Conf/list.xls"
                            else:
                                self.confpack['listpath'] = self.conf['dir'] + "Conf/list.xlsx"
                                
                        rep = self.askYesNo("Attention","Le package en cours de création/modification contient déjà un fichier list.\nConfirmez vous le nouvel import ?",fenconf)
                        if ( rep == "yes" ):
                            importcsv = True
                        else:
                            importcsv = False
                    else:
                        importcsv = True

                    if ( importcsv == True ):
                        CSV = tkFileDialog.askopenfilename(parent=fenconf,title="Choisir le CSV", initialdir=self.dir_path, \
                            initialfile="", filetypes = [("Fichiers CSV","*.csv"),("Fichiers Excel",("*.xls","*.xlsx"))]) 
                        EXT = CSV.split(".")[1]
                        shutil.copy2(CSV,self.conf['dir'] + "Conf/list."+EXT)
                        self.confpack['listpath'] = self.conf['dir'] + "Conf/list."+EXT
                        if ( EXT == "csv" or EXT == "CSV"):
                            self.confpack['listformat'] = "CSV"
                        else:
                            self.confpack['listformat'] = "XLS"
                    
                    self.ButtonPkgPlace['state'] = "disabled"
                    if ( CommandOpenPackage(path=self.conf['dir'],silence=1) ):
                        self.ButtonImportCSV['state'] = "active"
                        self.ButtonConfig['state'] = "active"
                        self.ButtonAspects['state'] = "active"
                        self.ButtonProperties['state'] = "active"
                        self.ButtonImportFiles['state'] = "active"
                        self.ButtonLoad['state'] = "active"
                    else:
                        self.ButtonImportCSV['state'] = "active"
                        self.ButtonConfig['state'] = "active"
                        self.ButtonAspects['state'] = "active"
                        self.ButtonProperties['state'] = "active"
                        self.ButtonImportFiles['state'] = "active"
                    self.logger("Le fichier CSV a été importé.","Success")
                    return True
                except Exception, e:
                    return False

            def ImportFiles():
                self.confpack = self.get_confpack()
                Place = tkFileDialog.askdirectory(parent=fenconf,title="Choisir le dossier contenant les documents", initialdir=self.dir_path)
                try:
                    sourceDir = Place+"/"
                    fileinfo = self.parse_csv()
                    
                    self.Bar['maximum'] = len(self.fileinfo)
                    val = 1
                    pval = 0
                    
                    for line in fileinfo:
                        if ( os.path.exists(sourceDir + fileinfo[line]['name']) == True):
                            shutil.copy2(sourceDir + fileinfo[line]['name'],self.conf['dir'] + "Orig/" + fileinfo[line]['name'])
                            pval = pval + val
                            self.Bar['value'] = pval
                            self.Treatment.update()
                        else:
                            self.logger("Document du CSV manquant dans le répertoire","Error")
                            return False
                    self.ButtonPkgPlace['state'] = "disabled"
                    if ( CommandOpenPackage(path=self.conf['dir'],silence=1) ):
                        self.ButtonImportCSV['state'] = "active"
                        self.ButtonConfig['state'] = "active"
                        self.ButtonAspects['state'] = "active"
                        self.ButtonProperties['state'] = "active"
                        self.ButtonImportFiles['state'] = "active"
                        self.ButtonLoad['state'] = "active"
                    else:
                        self.ButtonImportCSV['state'] = "active"
                        self.ButtonConfig['state'] = "active"
                        self.ButtonAspects['state'] = "active"
                        self.ButtonProperties['state'] = "active"
                        self.ButtonImportFiles['state'] = "active"
                    pval = pval + val
                    self.Bar['value'] = pval
                    self.logger("Les documents ont été importés.","Success")
                    return True
                except Exception, e:
                    self.logger(str(e),"Error")
                    return False
            
            def Config():
                def commandSave():
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
                        fields = self.get_fieldsCSV()
                        try:
                            file = open(self.conf['dir']+"Conf/package.conf", "wb")
                            file.write("[PACKAGE]\n")
                            file.write("OLDNAME="+str(self.getIndex(var['OLDNAME'].get(),fields))+"\n")
                            file.write("NEWNAME="+str(self.getIndex(var['NEWNAME'].get(),fields))+"\n")
                            file.write("TITLE="+str(self.getIndex(var['TITLE'].get(),fields))+"\n")
                            file.write("DESC="+str(self.getIndex(var['DESC'].get(),fields))+"\n")
                            file.write("TAGS="+str(self.getIndex(var['TAGS'].get(),fields))+"\n")
                            file.write("DESTPATH="+str(self.getIndex(var['DESTPATH'].get(),fields))+"\n")
                            file.close()
                        except Exception, e:
                            self.logger(str(e),"Error")
                        fenconfpkg.destroy()
                        self.ButtonPkgPlace['state'] = "disabled"
                        if ( CommandOpenPackage(path=self.conf['dir'],silence=1) ):
                            self.ButtonImportCSV['state'] = "active"
                            self.ButtonConfig['state'] = "active"
                            self.ButtonAspects['state'] = "active"
                            self.ButtonProperties['state'] = "active"
                            self.ButtonImportFiles['state'] = "active"
                            self.ButtonLoad['state'] = "active"
                        else:
                            self.ButtonImportCSV['state'] = "active"
                            self.ButtonConfig['state'] = "active"
                            self.ButtonAspects['state'] = "active"
                            self.ButtonProperties['state'] = "active"
                            self.ButtonImportFiles['state'] = "active"
                    else:
                        self.messageShow("Attention","Vous devez renseigner tous les champs.", fenconfpkg)
                    
                def updateList(evt, name, key, values):
                    value = name.entry.get()
                    
                    already = False
                    for k in values:
                        if ( k != name and value == values[k]):
                            already=True
                            
                    if ( already == False ):
                        values[key] = value
                    else:
                        self.messageShow("Attention","Champ déjà utilisé !!", fenconfpkg)
                        var[key].set("")
                    
                fenconfpkg = Toplevel(fenconf)
                fenconfpkg.title("Configuration (package.conf)")

                LabelKeys = Label(fenconfpkg, anchor=W, text="Clefs", width=20, font=titlefont)
                LabelFields = Label(fenconfpkg, anchor=W, text="Champs (CSV)", width=20, font=titlefont)

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
                
                if (os.path.exists(self.conf['dir'] + "Conf/package.conf") == True):
                    confpacktmp = self.get_confpack()
                    fields = self.get_fieldsCSV()
                    var['OLDNAME'].set(fields[confpacktmp['name']])
                    var['NEWNAME'].set(fields[confpacktmp['newname']])
                    var['TITLE'].set(fields[confpacktmp['title']])
                    var['DESC'].set(fields[confpacktmp['desc']])
                    var['TAGS'].set(fields[confpacktmp['tags']])
                    var['DESTPATH'].set(fields[confpacktmp['path']])

                LabelKeys.grid(row=0, column=0)
                LabelFields.grid(row=0, column=1)

                LabelOLDNAME.grid(row=1, column=0)
                LabelNEWNAME.grid(row=2, column=0)
                LabelTITLE.grid(row=3, column=0)
                LabelDESC.grid(row=4, column=0)
                LabelTAGS.grid(row=5, column=0)
                LabelDESTPATH.grid(row=6, column=0)

                listOLDNAME.grid(row=1, column=1)
                listNEWNAME.grid(row=2, column=1)
                listTITLE.grid(row=3, column=1)
                listDESC.grid(row=4, column=1)
                listTAGS.grid(row=5, column=1)
                listDESTPATH.grid(row=6, column=1)

                fenconfpkg.withdraw()
                l = self.posfen(fenconfpkg,bottom=True)
                self.displayBottom(fenconfpkg,l,7,partial(commandSave))
                return
            
            def ConfigAspects():
                def commandSave():
                    listaspects = Aspects.get('1.0', 'end')
                    tabaspects = listaspects.split("\n")
                    try:
                        file = open(self.conf['dir']+"Conf/aspects.conf", "wb")
                        for aspect in tabaspects:
                            if ( aspect != "" ):
                                file.write(aspect+"\n")
                        file.close()
                        fenconfasp.destroy()
                        self.ButtonPkgPlace['state'] = "disabled"
                        if ( CommandOpenPackage(path=self.conf['dir'],silence=1) ):
                            self.ButtonImportCSV['state'] = "active"
                            self.ButtonConfig['state'] = "active"
                            self.ButtonAspects['state'] = "active"
                            self.ButtonProperties['state'] = "active"
                            self.ButtonImportFiles['state'] = "active"
                            self.ButtonLoad['state'] = "active"
                        else:
                            self.ButtonImportCSV['state'] = "active"
                            self.ButtonConfig['state'] = "active"
                            self.ButtonAspects['state'] = "active"
                            self.ButtonProperties['state'] = "active"
                            self.ButtonImportFiles['state'] = "active"
                    except Exception, e:
                        self.logger(str(e),"Error")
                    return
                
                fenconfasp = Toplevel(fenconf)
                fenconfasp.title("Configutation des aspects")
                
                Aspects = Text(fenconfasp,bg="white",width=35,height=10)
                
                if (os.path.exists(self.conf['dir'] + "Conf/aspects.conf") == True):
                    file = open(self.conf['dir'] + "Conf/aspects.conf", "r")
                    line = 1
                    for aspect in file.readlines():
                        if ( aspect != "\n" ):
                            Aspects.insert(str(line)+".0",aspect)
                            line=line+1
                    file.close()
                
                Aspects.grid(sticky="W",row=0, columnspan=2)

                fenconfasp.withdraw()
                l = self.posfen(fenconfasp,bottom=True)
                self.displayBottom(fenconfasp,l,1,partial(commandSave))
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
                    
                    props[idx]['name'].grid(sticky="W",row=idx, column=0)
                    props[idx]['formatlist'].grid(sticky="W",row=idx, column=1)
                    if ( type == "STA" ):
                        props[idx]['value'].grid(sticky="W",row=idx, column=2)
                    else:
                        props[idx]['valuelist'].grid(sticky="W",row=idx, column=2)
                    Block.update_idletasks()
                    return
                            
                def commandSave():
                    file = open(self.conf['dir']+"Conf/properties.csv", "wb")
                    fields = self.get_fieldsCSV()
                    for idx in props:
                        name = props[idx]['name'].get()
                        format = props[idx]['format']
                        type = props[idx]['type']
                        if ( type == "STA" ):
                            value = props[idx]['value'].get().decode("utf-8").encode("iso8859_1")
                        else:
                            value = str(self.getIndex(props[idx]['value'],fields))
                        file.write(name+";"+format+";"+type+";"+value+"\n")
                    file.close()
                    fenconfpro.destroy()
                    self.ButtonPkgPlace['state'] = "disabled"
                    if ( CommandOpenPackage(path=self.conf['dir'],silence=1) ):
                        self.ButtonImportCSV['state'] = "active"
                        self.ButtonConfig['state'] = "active"
                        self.ButtonAspects['state'] = "active"
                        self.ButtonProperties['state'] = "active"
                        self.ButtonImportFiles['state'] = "active"
                        self.ButtonLoad['state'] = "active"
                    else:
                        self.ButtonImportCSV['state'] = "active"
                        self.ButtonConfig['state'] = "active"
                        self.ButtonAspects['state'] = "active"
                        self.ButtonProperties['state'] = "active"
                        self.ButtonImportFiles['state'] = "active"
                    return
                
                props = {}
                
                fenconfpro = Toplevel(fenconf)
                fenconfpro.title("Configuration des propriétés")
                
                AddSTA = Button(fenconfpro, command=addSta, text="Ajout prop. statique", relief=RAISED, font=buttonfont)
                AddDYN = Button(fenconfpro, command=addDyn, text="Ajout prop. dynamique", relief=RAISED, font=buttonfont)
                
                AddSTA.grid(sticky="W",row=0, column=0)
                AddDYN.grid(sticky="E",row=0, column=1)
                
                Block = Canvas(fenconfpro)

                Block.grid(sticky="E",row=1, columnspan=2)
                
                if (os.path.exists(self.conf['dir'] + "Conf/properties.csv") == True):
                    fields = self.get_fieldsCSV()
                    file = open(self.conf['dir'] + "Conf/properties.csv", "r")
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
                                                       
                            props[idx]['name'].grid(sticky="W",row=idx, column=0)
                            props[idx]['formatlist'].grid(sticky="W",row=idx, column=1)
                            
                            
                            if ( props[idx]['type'] == "STA" ):
                                props[idx]['value'] = Entry(Block,bg="white",width=15)
                                props[idx]['value'].insert(0,unicode(line[3],"iso8859_1"))
                                props[idx]['value'].grid(sticky="W",row=idx, column=2, padx=5)
                            else:
                                props[idx]['valuetix'] = Tix.StringVar()
                                props[idx]['valuelist'] = Tix.ComboBox(Block, editable=1, dropdown=1, variable=props[idx]['valuetix'], width=20, listwidth=50)
                                gen_listFields(props[idx]['valuelist'])
                                props[idx]['valuelist'].entry.config(width=15, state='readonly')
                                props[idx]['valuelist'].slistbox.listbox.bind('<ButtonRelease-1>',partial(updateList,name=props[idx]['valuelist'],key=idx,word="value"))
                                props[idx]['value'] = fields[int(line[3])]
                                props[idx]['valuetix'].set(fields[int(line[3])])
                                props[idx]['valuelist'].grid(sticky="W",row=idx, column=2)
                            idx=idx+1
                    file.close()
                
                fenconfpro.withdraw()
                l = self.posfen(fenconfpro,bottom=True)
                self.displayBottom(fenconfpro,l,19,partial(commandSave))
                
                return
            
            def loadPackage():
                fenconf.destroy()
                CommandOpenPackage(path=self.conf['dir'])
                return
            
            fenconf = Toplevel(fen)
       
            fenconf.title("Import Alfresco - Création d'un package")
            
            defaultfont = tkFont.Font(fenconf, size=10, family='Verdana', weight='bold')
            helpfont = tkFont.Font(fenconf, size=10, family='Verdana', slant='italic')
            buttonfont = tkFont.Font(fenconf, size=10, family='Verdana', weight='bold')
            
            PathCSV = Label(fenconf, fg=FGHIDDEN, width=60, font=defaultfont)
            Path = Label(fenconf, fg=FGHIDDEN, width=60, font=defaultfont)
            
            self.ButtonPkgPlace = Button(fenconf, text='Nouveau', command=PkgPlace, font=buttonfont, relief=GROOVE,width=30)
            self.ButtonImportCSV = Button(fenconf, text='Importer le CSV', command=ImportCSV, font=buttonfont, relief=GROOVE,width=30)
            self.ButtonImportCSV.config(state=DISABLED)
            self.ButtonImportFiles = Button(fenconf, text='Importer les documents', command=ImportFiles, font=buttonfont, relief=GROOVE,width=30)
            self.ButtonImportFiles.config(state=DISABLED)
            self.ButtonConfig = Button(fenconf, text='Gestion package.conf', command=Config, font=buttonfont, relief=GROOVE,width=30)
            self.ButtonConfig.config(state=DISABLED)
            self.ButtonAspects = Button(fenconf, text='Gestion aspects.conf', command=ConfigAspects, font=buttonfont, relief=GROOVE,width=30)
            self.ButtonAspects.config(state=DISABLED)
            self.ButtonProperties = Button(fenconf, text='Gestion properties.csv', command=ConfigProperties, font=buttonfont, relief=GROOVE,width=30)
            self.ButtonProperties.config(state=DISABLED)
            self.ButtonLoad = Button(fenconf, text='Exploiter', command=loadPackage, font=buttonfont, relief=GROOVE,width=30)
            self.ButtonLoad.config(state=DISABLED)
            
            if ( self.conf['open'] != "" ):
                self.ButtonPkgPlace['state'] = "disabled"
                self.ButtonConfig['state'] = "active"
                self.ButtonAspects['state'] = "active"
                self.ButtonProperties['state'] = "active"
                self.ButtonLoad['state'] = "active"
            
            self.ButtonPkgPlace.pack(side=TOP, anchor=W, expand=NO)
            self.ButtonImportCSV.pack(side=TOP, anchor=W, expand=NO)
            self.ButtonConfig.pack(side=TOP, anchor=W, expand=NO)
            self.ButtonAspects.pack(side=TOP, anchor=W, expand=NO)
            self.ButtonProperties.pack(side=TOP, anchor=W, expand=NO)
            self.ButtonImportFiles.pack(side=TOP, anchor=W, expand=NO)
            self.ButtonLoad.pack(side=TOP, anchor=W, expand=NO)
            
            fenconf.withdraw()
            self.posfen(fenconf)

            return
        
        def CommandClosePackage():
            self.conf['dir'] = ""
            self.conf['open'] = ""
            self.confpack = {}
            self.Logger.delete('1.0', END)
            self.ButtonGenerate['state'] = "disabled"
            self.ButtonTestHost['state'] = "disabled"
            self.ButtonUpload['state'] = "disabled"
            self.Force['state'] = "disabled"
            self.Bar['value'] = 0
            self.UpdateGuide("Etape 1 : Choisissez le dossier contenant le package")
        
        ### Fenêtre principale
        fen = Tix.Tk()

        BG = "#4F4F4F"
        FGHIDDEN = "#4F4F4F"
        FG = "#FFF"
        FGSTATE = "red"
        
        # Polices
        defaultfont = tkFont.Font(fen, size=10, family='Verdana', slant='italic')
        titlefont = tkFont.Font(fen, size=10, family='Verdana', weight='bold')
        menufont = tkFont.Font(fen, size=10, family='Verdana', weight='bold')
        submenufont = tkFont.Font(fen, size=10, family='Verdana')
        guidefont = tkFont.Font(fen, size=11, family='Verdana', weight='bold')
        buttonfont = tkFont.Font(fen, size=10, family='Verdana', weight='bold')
        forcefont = tkFont.Font(fen, size=10, family='Verdana', weight='bold')
        
        fen.title('Import Alfresco - Version '+self.version)
        fen.config(bg=BG, relief=GROOVE)

        FM = Frame(fen, bg="#eee")
        FM2 = Frame(fen, bg="#eee")
        FM3 = Frame(fen, bg="#eee")
        
        # Affichage des logs
        self.Logger = Text(fen, bg="#5f5f5f", fg="#ccc", width=110, height=40, highlightthickness="0")
        self.Logger.tag_configure('Error', foreground="#FF8888")
        self.Logger.tag_configure('Success', foreground="#88FF88")
        # Conteneur du traitement
        self.Treatment = Label(fen, bg=BG, fg=FGHIDDEN, width=60, font=defaultfont)
        
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
        self.ButtonGenerate = Button(FM, text='Générer le package', command=CommandGenerate, font=buttonfont, relief=GROOVE)
        self.ButtonGenerate.config( height = 1, width = 15 )
        self.ButtonGenerate.config(state=DISABLED)

        # Import dans Alfresco
        # Bouton d'upload
        self.ButtonUpload = Button(FM, text='Importer (mode Bulk Import)', command=CommandUpload, font=buttonfont, relief=GROOVE)
        self.ButtonUpload.config( height = 1, width = 15 )
        self.ButtonUpload.config(state=DISABLED)
        
        self.ButtonTestHost = Button(FM, text='Test connexions', command=CommandTestHost, font=buttonfont, relief=GROOVE)
        self.ButtonTestHost.config( height = 1, width = 15 )
        self.ButtonTestHost.config(state=DISABLED)
        
        self.var1 = IntVar()
        self.Force = Checkbutton(FM2, text = "Mise à jour uniquement (CMIS)", highlightthickness="0", font="forcefont", variable = self.var1 , command=ChangeMode)
        self.Force.config(state=DISABLED)

        # Affichage du guide
        self.Guide = Text(FM3, bg="ivory", fg="#222", width=110, height=1.4, font=guidefont)

        # Bar de progression
        s = ttk.Style()
        s.theme_use('clam')
        s.configure("red.Horizontal.TProgressbar", foreground='red', background='red')
        self.Bar = ttk.Progressbar(FM3, style="red.Horizontal.TProgressbar", orient="horizontal", length=300, mode="determinate", maximum=300)

        # Bouton de nettoyage de l'écran des logs
        #"Clear = Button(FM2, text="Nettoyer les logs", command=CommandClearLog, font=buttonfont, relief=RAISED)

        # Placements
        self.ButtonGenerate.pack(side=LEFT, anchor=W, fill=X, expand=YES)
        self.ButtonTestHost.pack(side=LEFT, anchor=W, fill=X, expand=YES)
        self.ButtonUpload.pack(side=LEFT, anchor=W, fill=X, expand=YES)
        FM.pack(fill=X)
        FM2.pack(fill=X)
        FM3.pack(fill=X)
        self.Logger.pack(side=BOTTOM, fill=X)
        self.Guide.pack(side=BOTTOM, fill=X)
        self.Bar.pack(side=BOTTOM, fill=X)
        #self.Mode.pack(side=RIGHT, fill=X)
        self.Force.pack(side=RIGHT, fill=X)
        #Clear.pack(side=RIGHT, fill=X)

        #Clear.pack(side=BOTTOM, anchor=W, fill=X, expand=YES)

        self.Guide.tag_configure('tag-center', justify='center')
        self.UpdateGuide("Etape 1 : Choisissez le dossier contenant le package")

        #Center(fen)
        fen.config(menu=menubar)
        
        # Taille et position de la fenêtre
        fen.resizable(width=False, height=False)
        
        self.posfen(fen, 1024, 660)
        
        fen.mainloop()

        return

    # Gui Composer
    
    # Taille et position de la fenêtre
    def posfen(self,fen, FENW=0, FENH=0, bottom=False):
        RESX = fen.winfo_screenwidth()
        if ( RESX > 1920 ):
            RESX = 1920
        RESY = fen.winfo_screenheight()
        
        if ( FENW == 0 ):
            fen.after(500,fen.update_idletasks())
            FENW = fen.winfo_reqwidth()
            FENH = fen.winfo_reqheight()
            if ( bottom ):
                FENH=FENH+50
                
        POSX = (RESX - FENW) /2
        POSY = (RESY - FENH) /2

        fen.geometry(str(FENW)+"x"+str(FENH)+"+"+str(POSX)+"+"+str(POSY))
        fen.deiconify()
        
        return FENW

    def displayBottom(self, fen, l, r, commandSave):
        buttonfont = tkFont.Font(fen, size=10, family='Verdana', weight='bold')
        
        Line = Canvas(fen, width=l, height=10, highlightthickness=0)
        Line.grid(sticky="W",row=r, columnspan=2)
        Line.create_line(0,4,1000,4, fill="#555")
        Quit = Button(fen, text="Quitter", command=fen.destroy, relief=RAISED, font=buttonfont)
        Save = Button(fen, text="Sauver", command=commandSave, relief=RAISED, font=buttonfont)

        Save.grid(sticky="W",row=r+1, column=0, padx=5)
        Quit.grid(sticky="E",row=r+1, column=1, padx=5)
        
    def testWkspace(self, path):
        try:
            client = CmisClient(self.conf['url'], self.conf['user'], self.conf['password'])
            repo = client.defaultRepository

            Folder = repo.getObjectByPath(path)
            return [True,path + "' : OK", "Success"]
        except Exception, e:
            return [False,path + "' : Introuvable", ""]

    def testCMIS(self):
        try:
            client = CmisClient(self.conf['url'], self.conf['user'], self.conf['password'])

            repo = client.defaultRepository
            REQ = "select * from cmis:folder where cmis:name = 'Sites'"

            results = repo.query(REQ)
            
            if (len(results) != 0):
                return str(results[0].id)
            else:
                return False
        except Exception, e:
            self.logger(str(e),"Error")
            return False
        
    def testBulkImport(self):
        try:
            rh = RESTHelper()
            rh.login(self.conf['user'], self.conf['password'], self.conf['host'], 8080)
            data = rh.statusbulkimport()
            return True
        except Exception, e:
            self.logger(str(e),"Error")
            return False
    
    def getStatusBulkImport(self):
        try:
            rh = RESTHelper()
            rh.login(self.conf['user'], self.conf['password'], self.conf['host'], 8080)
            data = rh.statusbulkimport()
            return json.load(data)
        except Exception, e:
            self.logger(str(e),"Error")
            return False
    
    def OpenHost(self, mode, silence):
        self.Logger.delete('1.0', END)
        if ( self.conf['host'] != "" ):
            self.conf['url'] = self.conf['urltemp'].replace("__HOST__", self.conf['host'])
            
            testbulkimport = self.testBulkImport()

            if ( testbulkimport ):
                self.logger("Test connexion Bulk Import Tool : OK","Success",silence)
            else:
                self.logger("Test connexion Bulk Import Tool : Error","Error",silence)
                return False
            
            if (os.path.exists(self.conf['dir'] + "Conf/package.conf")):
                testhost = self.testCMIS()
                
                if (testhost != False ):
                    self.logger("Test connexion CMIS Alfresco ("+self.conf['host']+") : OK","Success",silence)
                    self.logger("","",silence)
                    self.logger("Test des destinations du CSV :","",silence)
                    self.logger("","",silence)

                    pathlist = {}
                    WKTEST = True
                    for fid in self.fileinfo:

                        path = self.fileinfo[fid]['upath']

                        if ( path not in pathlist ):
                            pathlist[path] = True
                            test = self.testWkspace(path)
                            self.logger(test[1], test[2],silence)
                            if ( test[0] == False ):
                                WKTEST = False
                else:
                    WKTEST = False

                if ( WKTEST == True ):
                    if ( mode == "BULKIMPORTTOOL" ):
                        self.ButtonUpload['state'] = "active"
                        self.Force['state'] = "active"
                    else:
                        self.ButtonUpload['state'] = "active"
                        self.var1.set(1)
                        self.Force['state'] = "active"
                    return True
                else:
                    self.Force['state'] = "disabled"
                    self.var1.set(0)
                    self.logger("","",silence)
                    self.logger("Test des destinations : dossiers manquants (sans conséquences en mode Bulk Import Tool)","",silence)
                    self.logger("","",silence)
                    self.logger("Pour le mode CMIS (mise à jour), le mode Bulk Import Tool doit être lancé en premier)","",silence)
                    self.ButtonUpload['text'] = "Importer (mode Bulk Import)"
                    if ( mode == "BULKIMPORTTOOL" ):
                        return True
                    else:
                        return False
        else:
            self.logger("","",silence)
            self.logger("Le serveur hôte Alfresco n'est pas configuré","Error",silence)
            

    # Get files informations from CSV
#    def parse_csv(self):
#        self.fileinfo = {}
#
#        csvfile = open(self.conf['dir'] + 'Conf/list.csv', "rb")
#        reader = csv.reader(csvfile, delimiter=';', quotechar='"')
#        firstline = 1    
#        for row in reader:
#            if (firstline == 1):
#                firstline = 0
#            else:
#                fid = str(row[0])
#                name = unicode(row[self.confpack['name']], "iso8859_1")
#                newname = unicode(row[self.confpack['newname']], "iso8859_1")
#                bname = row[self.confpack['name']]
#                bnewname = row[self.confpack['newname']]
#                title = unicode(row[self.confpack['title']], "iso8859_1")
#                description = unicode(row[self.confpack['desc']], "iso8859_1")
#                tags = unicode(row[self.confpack['tags']], "iso8859_1").split(",")
#                path = unicode(row[self.confpack['path']], "iso8859_1")
#                upath = row[self.confpack['path']].decode('iso8859_1').encode('utf-8')
#
#                idx = "F" + fid
#
#                self.fileinfo[idx] = {}
#                self.fileinfo[idx]['id'] = fid
#                self.fileinfo[idx]['title'] = title
#                self.fileinfo[idx]['description'] = description
#                self.fileinfo[idx]['tags'] = tags
#                self.fileinfo[idx]['name'] = name
#                self.fileinfo[idx]['newname'] = newname
#
#                # Version sans unicode
#                self.fileinfo[idx]['bname'] = bname
#                self.fileinfo[idx]['bnewname'] = bnewname
#
#                self.fileinfo[idx]['path'] = path
#                self.fileinfo[idx]['upath'] = upath
#
#                self.fileinfo[idx]['properties'] = {}
#
#                self.fileinfo[idx]['properties'] = self.get_properties(self.fileinfo[idx]['properties'], row)
#
#        csvfile.close()
#
#        return self.fileinfo

    def get_fieldsCSV(self):
        fields = {}
        
        if ( self.confpack['listformat'] == "XLS" ):
            wb = xlrd.open_workbook(self.confpack['listpath'])
        else:   
            csvfile = open(self.confpack['listpath'], "rb")
            reader = csv.reader(csvfile, delimiter=';', quotechar='"')
        
        if ( self.confpack['listformat'] == "XLS" ):
            reader = []
            feuils =  wb.sheet_names()
            sh = wb.sheet_by_name(feuils[0])
            for rownum in range(sh.nrows):
                row = sh.row_values(rownum)
                reader.append(row)
        
        for rows in reader:
            i=0
            for row in rows:
                if ( self.confpack['listformat'] == "CSV" ):
                    fields[i] = unicode(row,"iso8859_1")
                else:
                    fields[i] = row
                i=i+1
            break
        
        if ( self.confpack['listformat'] == "CSV" ):
            csvfile.close()
        
        return fields
    
        # Get files informations from XLS or CSV
    def parse_csv(self):
        self.fileinfo = {}
        
        self.confpack['listformat'] = "XLS"
        
        if ( os.path.exists(self.conf['dir'] + 'Conf/list.xls') == True ):
            wb = xlrd.open_workbook(self.conf['dir'] + 'Conf/list.xls')
            self.confpack['listpath'] = self.conf['dir'] + 'Conf/list.xls'
        elif ( os.path.exists(self.conf['dir'] + 'Conf/list.xlsx') == True ):
            wb = xlrd.open_workbook(self.conf['dir'] + 'Conf/list.xlsx')
            self.confpack['listpath'] = self.conf['dir'] + 'Conf/list.xlsx'
        elif ( os.path.exists(self.conf['dir'] + 'Conf/list.csv') == True ):    
            csvfile = open(self.conf['dir'] + 'Conf/list.csv', "rb")
            reader = csv.reader(csvfile, delimiter=';', quotechar='"')
            self.confpack['listpath'] = self.conf['dir'] + 'Conf/list.csv'
            self.confpack['listformat'] = "CSV"
        else:
            return self.fileinfo
        
        firstline = 1
        
        if ( self.confpack['listformat'] == "XLS" ):
            reader = []
            feuils =  wb.sheet_names()
            sh = wb.sheet_by_name(feuils[0])
            for rownum in range(sh.nrows):
                row = sh.row_values(rownum)
                reader.append(row)

        for row in reader:
            if (firstline == 1):
                firstline = 0
            else:
                if ( isinstance(row[0],float) ):
                    fid = "%s" % int(row[0])
                else:
                    fid = str(row[0])
                if ( self.confpack['listformat'] == "CSV"):
                    name = unicode(row[self.confpack['name']], "iso8859_1")
                    newname = unicode(row[self.confpack['newname']], "iso8859_1")
                    bname = row[self.confpack['name']]
                    bnewname = row[self.confpack['newname']]
                    title = unicode(row[self.confpack['title']], "iso8859_1")
                    description = unicode(row[self.confpack['desc']], "iso8859_1")
                    tags = unicode(row[self.confpack['tags']], "iso8859_1").split(",")
                    path = unicode(row[self.confpack['path']], "iso8859_1")
                    upath = row[self.confpack['path']].decode('iso8859_1').encode('utf-8')
                else:
                    name = row[self.confpack['name']]
                    newname = row[self.confpack['newname']]
                    bname = row[self.confpack['name']]
                    bnewname = row[self.confpack['newname']]
                    title = row[self.confpack['title']]
                    description = row[self.confpack['desc']]
                    tags = row[self.confpack['tags']].split(",")
                    path = row[self.confpack['path']]
                    upath = row[self.confpack['path']].encode('utf-8')

                idx = "F" + fid

                self.fileinfo[idx] = {}
                self.fileinfo[idx]['id'] = fid
                self.fileinfo[idx]['title'] = title
                self.fileinfo[idx]['description'] = description
                self.fileinfo[idx]['tags'] = tags
                self.fileinfo[idx]['name'] = name
                self.fileinfo[idx]['newname'] = newname
                
                # Version sans unicode
                self.fileinfo[idx]['bname'] = bname
                self.fileinfo[idx]['bnewname'] = bnewname

                self.fileinfo[idx]['path'] = path
                self.fileinfo[idx]['upath'] = upath
                
                self.fileinfo[idx]['properties'] = {}

                self.fileinfo[idx]['properties'] = self.get_properties(self.fileinfo[idx]['properties'], row)

        if ( self.confpack['listformat'] == "CSV" ):
            csvfile.close()
            
        return self.fileinfo

    def get_aspects(self):
        file = open(self.conf['dir'] + "Conf/aspects.conf", "r")
        Aspects = []
        for aspect in file.readlines():
            if ( aspect != "\n"):
                Aspects.append(aspect.rstrip())
        file.close()
        return Aspects

    def get_properties(self,tab, data):
        if (os.path.exists(self.conf['dir'] + "Conf/properties.csv") == True):
            csvfile = open(self.conf['dir'] + 'Conf/properties.csv', "rb")
            reader = csv.reader(csvfile, delimiter=';', quotechar='"')

            for row in reader:
                if (row[2] == "DYN"):
                    if ( self.confpack['listformat'] == "CSV" ):
                        tab[row[0]] = data[int(row[3])]
                    else:
                        if ( isinstance(data[int(row[3])],float) and row[1] == "TXT"):
                            data[int(row[3])] = int(data[int(row[3])])
                        tab[row[0]] = u"%s" % data[int(row[3])]
                else:
                    if ( self.confpack['listformat'] == "XLS" ):
                        tab[row[0]] = row[3].decode("iso8859_1")

                if (row[1] == "DATE"):
                    sdate = tab[row[0]].split("/")
                    tab[row[0]] = sdate[2] + "-" + sdate[1] + "-" + sdate[0] + "T00:00:00.000+00:00"
                if (row[1] == "TXT"):
                    if ( self.confpack['listformat'] == "CSV" ):
                        tab[row[0]] = unicode(tab[row[0]], "iso8859_1")
                if (row[1] == "NUM"):
                    tab[row[0]] = int(tab[row[0]])

            csvfile.close()
            return tab
        else:
            return {}

    def generatepack(self):  
        counter = 1
        self.Bar['maximum'] = len(self.fileinfo)
        val = 1
        pval = 0
        
            
        for fid in self.fileinfo:
            name = self.fileinfo[fid]['name']
            newname = self.fileinfo[fid]['newname']
            title = self.fileinfo[fid]['title']
            description = self.fileinfo[fid]['description']
            path = self.fileinfo[fid]['path']
            
            if ( os.path.exists(self.conf['dir'] + self.confpack['PKGID'] + path +'/'+newname) ):
                self.logger(str(counter) + " - " + newname + " existe déjà","")
                
                pval = pval + val
                self.Bar['value'] = pval

                self.logger(str(counter) + " - " + newname,"")
                counter = counter + 1
            else :
                pdf = PdfFileReader(open(self.conf['dir'] + "Orig/" + name, 'rb'))

                if (pdf.getIsEncrypted()):
                    pdf.decrypt('')

                data = {u'/Title':u'%s' % title, u'/Subject':u'%s' % description}

                self.add_metadata(name, path, newname, data)
                LOG = u"Traitement des métadonnées du document '%s'" % (newname)

                self.genXMLProperties(fid)

                pval = pval + val
                self.Bar['value'] = pval

                self.logger(str(counter) + " - " + newname,"")
                counter = counter + 1

        pval = pval + val        
        self.Bar['value'] = pval

        if ( self.Bar['value'] >= self.Bar['maximum'] ):
            return True
        else:
            return False
        
    def genXMLProperties(self,fid):
        id = self.fileinfo[fid]['id']

        xmlfile = open(self.conf['dir'] + self.confpack['PKGID'] +self.fileinfo[fid]['path']+'/'+self.fileinfo[fid]['newname']+'.metadata.properties.xml', "wb")
        
        xmlfile.write('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE properties SYSTEM "http://java.sun.com/dtd/properties.dtd">\n<properties>\n')
        xmlfile.write('<entry key="separator"> # </entry>\n')
        xmlfile.write('<entry key="type">cm:content</entry>\n')
        
        Aspects = self.get_aspects()
        
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
        
        for prop in self.fileinfo[fid]['properties']:
            propname = prop
            propvalue = self.fileinfo[fid]['properties'][prop]

            if ( self.confpack['listformat'] == "CSV" ):
                xmlfile.write(u'<entry key="%s">%s</entry>\n' % (propname, propvalue))
            else:
                xmlfile.write('<entry key="%s">%s</entry>\n' % (propname, propvalue))
            
        xmlfile.write(u'<entry key="ialf:packageid">%s</entry>\n' % (self.confpack['PKGID']+"_"+id))
        
        xmlfile.write('</properties>\n')
        
        xmlfile.close()
        
        return

    def add_metadata(self,name, path, newname, data):
        merger = PdfFileMerger()

        with open(self.conf['dir'] + 'Orig/%s' % name, 'rb') as f0:
            merger.append(f0)

        merger.addMetadata(data)

        if not os.path.exists(os.path.dirname(self.conf['dir'] + self.confpack['PKGID'] + '%s/%s' % (path,newname))):
            os.makedirs(os.path.dirname(self.conf['dir'] + self.confpack['PKGID'] + '%s/%s' % (path,newname)))

        with open(self.conf['dir'] + self.confpack['PKGID'] + '%s/%s' % (path,newname), 'wb') as f1:
            merger.write(f1)

    def upload(self, mode):  
        counter = 1
        self.Bar['maximum'] = len(self.fileinfo)
        self.Bar['value'] = 0
        self.logger("","")
        self.logger("Début de l'import vers "+self.conf['host'],"")
        FileMissing = False
        for fid in self.fileinfo:
            newname = self.fileinfo[fid]['newname']
            #self.cmisCreate(newname, self.fileinfo[fid], force, counter)
            RETURN = self.updateCmis(self.fileinfo[fid], counter)
            if ( RETURN == False ):
                self.logger("Document '"+newname+"' manquant","Error")
                FileMissing = True
            counter=counter+1

        if ( self.Bar['value'] >= self.Bar['maximum'] ):
            if ( FileMissing ):
                self.logger("","")
                self.logger("ATTENTION: des documents du package sont manquant dans Alfresco. Vous devez sans doute relancer l'import en mode Bulk Import Tool.","Error")
                return "Missing"
            return True
        else:
            return False

    def getSitesId(self):
        try:
            client = CmisClient(self.conf['url'], self.conf['user'], self.conf['password'])

            repo = client.defaultRepository
            REQ = "select * from cmis:folder where cmis:name = 'Sites'"

            results = repo.query(REQ)

            if (len(results) != 0):   
                return str(results[0].id)
            else:
                return False
        except Exception, e:
            self.logger(str(e),"Error")
            return False

    def cmisCreate(self, name, onefileinfo, force, counter):
            id = onefileinfo['id']
            pathdst = onefileinfo['path']

            client = CmisClient(self.conf['url'], self.conf['user'], self.conf['password'])

            repo = client.defaultRepository

            Folder = repo.getObject("workspace://SpaceStore/" + self.conf['wkimports'])

            REQ = "select * from ialf:package where ialf:packageid = '" + self.confpack['PKGID'] + "_" + id + "'"

            results = repo.query(REQ)

             # Création du document
            if (len(results) == 0):
                try:
                    LOG = u"Création du document %s" % name
                    File = open(self.conf['dir'] + self.confpack['PKGID'] + pathdst + "/" + name, 'r')
                    Doc = Folder.createDocument(self.confpack['PKGID'] + "_" + id+".pdf", contentFile=File)
                    File.close()
                    self.logger(LOG,"")
                    REQ = "select * from cmis:document where cmis:name = '" + self.confpack['PKGID'] + "_" + id+".pdf" + "'"
                    force = 1
                except Exception, e:
                    self.logger(u"Problème de création du document %s" % name,"Error")
                    self.logger(str(e),"Error")
            else:
                LOG = u"%s déjà existant" % name
                self.logger(LOG,"")

            if (force == 1):
               self.updateCmis(onefileinfo, counter)

            self.Bar['value'] = counter
            
    def updateCmis(self, onefileinfo, counter):
        id = onefileinfo['id']

        client = CmisClient(self.conf['url'], self.conf['user'], self.conf['password'])

        repo = client.defaultRepository

        REQ = "select * from ialf:package where ialf:packageid = '" + self.confpack['PKGID'] + "_" + id + "'"

        results = repo.query(REQ)
        
        if (len(results) == 0):
            return False
        
        # Aspects
        if (len(results) != 0):
            try:
                LOG = u"--> Application des aspects du document %s" % onefileinfo['newname']
                objectId = str(results[0].id)
                Doc = repo.getObject("workspace://SpaceStore/" + objectId)
                Aspects = self.get_aspects()
                for aspect in Aspects:
                    Doc.addAspect("P:" + aspect)
                Doc.addAspect("P:ialf:package")
                self.logger(LOG,"")
            except Exception, e:
                self.logger(u"Problème d'ajout des aspects pour %s" % onefileinfo['newname'],"Error")
                self.logger(str(e),"Error")
                return False

        # Propriétés
        results = repo.query(REQ)
        if (len(results) != 0):
            try:
                LOG =  u"--> Création des propriétés du document %s" % onefileinfo['newname']
                objectId = str(results[0].id)
                Doc = repo.getObject("workspace://SpaceStore/" + objectId)
                props = {}

                self.logger(LOG,"")

                for prop in onefileinfo['properties']:
                    propname = prop
                    propvalue = onefileinfo['properties'][prop]

                    props[propname] = propvalue

                props["ialf:packageid"] = self.confpack['PKGID']+"_"+id
                props["cmis:name"] = onefileinfo['newname']
                Doc.updateProperties(props)
            except Exception, e:
                self.logger(u"Problème d'ajout des propriétés pour %s" % onefileinfo['newname'],"Error")
                self.logger(str(e),"Error")
                return False

        if (len(results) != 0):
            objectId = str(results[0].id)
            try:
                for tag in onefileinfo['tags']:
                    if (tag != ""):
                        split_objId = objectId.split(";")
                        if ( self.confpack['listformat'] == "CSV" ):
                            self.add_tags(split_objId[0], tag)
                        else:
                            self.add_tags(split_objId[0], tag.decode("utf-8").encode("utf-8"))
                            
                self.logger(u"--> Ajout des tags pour %s" % onefileinfo['newname'],"")
            except Exception, e:
                self.logger(u"Problème d'ajout des tags pour %s" % onefileinfo['newname'],"Error")
                self.logger(str(e),"Error")
                return False
        
        # Déplacement du document
#        try:
#            REQ = "select * from ialf:package where ialf:packageid = '" + self.confpack['PKGID'] + "_" + id + "'"
#            results = repo.query(REQ)
#            if (len(results) != 0):
#                objectId = str(results[0].id)
#                Doc = repo.getObject("workspace://SpaceStore/" + objectId)
#                parent = Doc.getObjectParents().getResults()[0]
#                if (str(parent) == self.conf['wkimports']):
#                    src = repo.getObject("workspace://SpaceStore/" + self.conf['wkimports'])
#                    dst = repo.getObject("workspace://SpaceStore/" + wkspacedst)
#                    Doc.move(src, dst)
#        except Exception, e:
#                self.logger(u"Problème de déplacement de %s" % name,"Error")
#                self.logger(str(e),"Error")
#                return False

        self.Bar['value'] = counter
            
    def transfertPack(self):
        import scpclient
        try:
            client = paramiko.SSHClient()
            client.load_system_host_keys()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(self.conf['host'], 22, self.conf['user_ssh'], self.conf['password_ssh'])
            
            createdir = False
            try:
#                with closing(scpclient.Read(client.get_transport(), self.conf['importdir']+"/"+self.confpack['PKGID'])) as scp:
#                    scp.receive('Sites.metadata.properties.xml')
                distdir = self.conf['importdir']+"/"+self.confpack['PKGID']
                if ( distdir != "/"):
                    client.exec_command("if [ -d '"+distdir+"' ]; then rm -fr "+distdir+" ; fi", timeout=60)
                    createdir=True
                else:
                    self.logger("Suppression du package distant impossible","Error")
                    return False
            except:
                createdir = True
            
            if ( createdir ):
                dirempty=self.conf['dir']+"EMPTY"
                if (os.path.exists(dirempty) == False):
                    os.mkdir(dirempty)
                with closing(scpclient.WriteDir(client.get_transport(), self.conf['importdir']+"/"+self.confpack['PKGID'])) as scp:
                    scp.send_dir(dirempty, override_mode=True, preserve_times=True)
                os.rmdir(dirempty)
                
#                with closing(scpclient.Write(client.get_transport(), self.conf['importdir']+"/"+self.confpack['PKGID'])) as scp:
#                    scp.send_file(self.conf['dir'] + self.confpack['PKGID'] + "/Sites.metadata.properties.xml", remote_filename="Sites.metadata.properties.xml")
            
            pathdir = self.conf['dir'] + self.confpack['PKGID'] + "/Sites"
            with closing(scpclient.WriteDir(client.get_transport(), self.conf['importdir']+"/"+self.confpack['PKGID'])) as scp:
                scp.send_dir(pathdir, override_mode=True, preserve_times=True)
                
            client.close()
            return True
        except Exception, e:
            self.logger(str(e),"Error")
            return False
        
    def returnStatus(self, result):
        tag=""
        self.logger("","")
        self.logger(" - Package                       : "+str(result['sourceParameters']['Source Directory']),"")
        
        status = "En attente"
            
        if ( str(result['processingState']) == "Failed" ):
            status = "Echoué"
            tag="Error"

        if ( str(result['processingState']) == "Succeeded" ):
            status = "Succès"
            tag="Success"
        
        self.logger(" - Statut                        : "+status,tag)
        self.logger(" - Durée du traitement           : "+str(result['duration']),"")
        self.logger(" - Taille doc. importés (octets) : "+str(result['targetCounters']['Bytes imported']['Count']),"")
        
        if ( result['inProgress'] == False ):
            progress = "Terminé"
        else:
            progress = "En cours"
        
        self.logger(" - Progression                   : "+progress,tag)
        self.logger("","")
            
        return result['inProgress']
    
    def initiateBulkImport(self):
        self.logger("","")
        self.logger("Transfert du package sur "+self.conf['host']+" (Bulk Import Tool). Patientez...","")
        self.logger("","")
        if ( self.transfertPack() ):
            self.logger("Transfert du package sur "+self.conf['host']+" OK","Success")
            self.logger("","")
            
            rh = RESTHelper()
            rh.login(self.conf['user'], self.conf['password'], self.conf['host'], 8080)

            try:
                rh.initiateBulkImport(self.conf['importdir']+"/"+self.confpack['PKGID']+"/","/")
                self.logger("Initialisation de l'import (Bulk Import Tool)","Success")
                
                result = self.getStatusBulkImport()
                while ( self.returnStatus(result) ):
                    time.sleep(2)
                    self.Logger.delete('7.0', END)
                    self.Logger.insert(END,"\n")
                    result = self.getStatusBulkImport()
             
                return True
            except Exception, e:
                self.logger(str(e),"Error")
                return False
        else:
            self.logger("Transfert du package sur "+self.conf['host']+" Error","Error")
            self.logger("","")

    def add_tags(self, objectId, tag):
        rh = RESTHelper()
        rh.login(self.conf['user'], self.conf['password'], self.conf['host'], 8080)

        rh.addTag("workspace", "SpacesStore", objectId, tag)

        rh.logout
        return


