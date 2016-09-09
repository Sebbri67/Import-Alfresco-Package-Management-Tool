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
import ttk
import uuid

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

    # Get Global Config
    def get_conf(This):
        config = ConfigParser.ConfigParser()
        
        file = open(This.dir_path+"/importalf.conf", "rb")

        config.readfp(file)
        
        file.close()
        
        #This.Logger.delete('1.0', END)
        
        #LOG = "ATTENTION : la configuration globale est incomplète : "
        
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
        This.confpack['wkspace'] = int(config.get("PACKAGE", "DESTWSP"))
        This.confpack['path'] = int(config.get("PACKAGE", "DESTPATH"))
        return This.confpack

    # Logger
    def logger(This, txt, tag):
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
    def posfen(This,fen, FENW, FENH):
        RESX = fen.winfo_screenwidth()
        if ( RESX > 1920 ):
            RESX = 1920
        RESY = fen.winfo_screenheight()
         
        POSX = (RESX - FENW) /2
        POSY = (RESY - FENH) /2
        
        fen.geometry(str(FENW)+"x"+str(FENH)+"+"+str(POSX)+"+"+str(POSY))
    
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

        def CommandOpenPackage():
            This.Logger.delete('1.0', END)
            
            # Dialog Box for choose dir package
            PackageDir = tkFileDialog.askdirectory(initialdir="/home/sb/Documents/DSI/GED/", title="Choisir le répertoire")
            PathDir['text'] = PackageDir
            This.conf['dir'] = PathDir['text'] + "/"

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

                csvfile = open(This.conf['dir'] + 'Conf/list.csv', "rb")
                reader = csv.reader(csvfile, delimiter=';', quotechar='"')

                fields = {}

                for row in reader:
                    fields = row
                    break

                csvfile.close()

                name = unicode(fields[This.confpack['name']], "iso8859_1")
                newname = unicode(fields[This.confpack['newname']], "iso8859_1")
                title = unicode(fields[This.confpack['title']], "iso8859_1")
                description = unicode(fields[This.confpack['desc']], "iso8859_1")
                tags = unicode(fields[This.confpack['tags']], "iso8859_1")
                wkspace = unicode(fields[This.confpack['wkspace']], "iso8859_1")
                path = unicode(fields[This.confpack['path']], "iso8859_1")

                This.logger("\nVérification correspondance des champs :\n","")
                This.logger("  Attributs   | Champs CSV","")
                This.logger("  ____________|___________    ","")
                This.logger("  OLDNAME     | "+name,"")
                This.logger("  NEWNAME     | "+newname,"")
                This.logger("  TITLE       | "+title,"")
                This.logger("  DESC        | "+description,"")
                This.logger("  TAGS        | "+tags,"")
                This.logger("  DESTPATH    | "+path,"")
                This.logger("  DESTWKS     | "+wkspace,"")

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
                        valuedetail = " : Valeur dynamique du champs CSV '"+unicode(fields[int(field)], "iso8859_1")+"'"

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
            else:
                This.logger("\n"+This.conf['dir'] + " : Erreur.", "Error")
                This.ButtonGenerate['state'] = "disabled"
                This.ButtonTestHost['state'] = "disabled"
                This.Host['state'] = "disabled"
                This.ButtonUpload['state'] = "disabled"

        def CommandUpload():
            mode = "BULKIMPORTTOOL"
            
            if ( This.var1.get() == 1 ):
                mode = "CMIS"
            
            TestHost = This.OpenHost(mode)
            
            if ( TestHost ):
                This.conf['wkimports'] = This.getImportId()

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
            TestHost = This.OpenHost(mode)
        
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
                #clearHosts()
                #listhosts = This.conf['hosts'].split(",")
                #count=0
                #for host in listhosts:
                #    This.Host.insert(count, host)
                #    count=count+1

            fenconf = Tix.Tk()
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
            This.posfen(fenconf,750,300)
            return
        
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
        menubar.add_cascade(label="Package", menu=package, font=menufont)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Configuration", command=CommandConfGlobale, font=submenufont)
        filemenu.add_command(label="Quitter", command=fen.destroy, font=submenufont)
        menubar.add_cascade(label="Général", menu=filemenu, font=menufont)


        # Choix du package
        # Bouton du package
        #ButtonPackage = Button(FM, text='Choisir le package', command=CommandOpenPackage, font=buttonfont, relief=RAISED)
        # Conteneur du package
        PathDir = Label(fen, bg=BG, fg=FGHIDDEN, width=60, font=defaultfont)

        # Traitement des documents
        # Bouton du traitement
        This.ButtonGenerate = Button(FM, text='Générer le package', command=CommandGenerate, font=buttonfont, relief=RAISED)
        This.ButtonGenerate.config(state=DISABLED)

        # Choix du host
        # Menu de choix du host
        #HostSelec = Tix.StringVar()  
        #This.Host = Tix.ComboBox(FM, editable=0, dropdown=1, bg="white", state=DISABLED, variable=HostSelec)
        
        #listhosts = This.conf['hosts'].split(",")
        #count=0
        #for host in listhosts:
        #    This.Host.insert(count, host)
        #    count=count+1

        #This.Host.slistbox.listbox.bind('<ButtonRelease-1>', CommandOpenHost)
        
        # Import dans Alfresco
        # Bouton d'upload
        This.ButtonUpload = Button(FM, text='Importer (mode Bulk Import)', command=CommandUpload, font=buttonfont, relief=RAISED)
        This.ButtonUpload.config(state=DISABLED)
        
        This.ButtonTestHost = Button(FM, text='Test connexions', command=CommandTestHost, font=buttonfont, relief=RAISED)
        This.ButtonTestHost.config(state=DISABLED)
        
        This.var1 = IntVar()
        This.Force = Checkbutton(FM2, text = "Mise à jour uniquement (CMIS)", highlightthickness="0", font="forcefont", variable = This.var1 , command=ChangeMode)
        This.Force.config(state=DISABLED)
        
        #var2 = IntVar()
        #This.Mode = Checkbutton(FM2, text = "Mode Bulk Import Tool", highlightthickness="0", font="forcefont", variable = var2)

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
        #This.Host.pack(side=LEFT, anchor=W,fill=X, expand=YES)
        #ButtonPackage.pack(side=LEFT, anchor=W, fill=X, expand=YES)
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
        This.posfen(fen, 1024, 620)
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
            return True
        except Exception, e:
            This.logger(str(e),"Error")
            return False
    
    def OpenHost(This, mode):
        This.Logger.delete('1.0', END)
        if ( This.conf['host'] != "" ):
            This.conf['url'] = This.conf['urltemp'].replace("__HOST__", This.conf['host'])
            LOG = "Test connexion Bulk Import Tool ("+This.conf['host']+")"
            testbulkimport = This.testBulkImport()
            
            if ( testbulkimport ):
                This.logger("Test connexion Bulk Import Tool : OK","Success")
            else:
                This.logger("Test connexion Bulk Import Tool : Error","Error")
                return False
            
            if (os.path.exists(This.conf['dir'] + "Conf/package.conf")):
                
                LOG = "Test connexion CMIS Alfresco ("+This.conf['host']+")"
                testhost = This.testCMIS()
                
                if (testhost != False ):
                    This.logger(LOG+" : OK","Success")
                    This.logger("\nTest des destinations du CSV :\n","")

                    pathlist = {}
                    WKTEST = True
                    for fid in This.fileinfo:

                        path = This.fileinfo[fid]['upath']

                        if ( path not in pathlist ):
                            pathlist[path] = True
                            test = This.testWkspace(path)
                            This.logger(test[1], test[2])
                            if ( test[0] == False ):
                                WKTEST = False
                else:
                    WKTEST = False

                if ( WKTEST == True ):
                    if ( mode == "BULKIMPORTTOOL" ):
                        This.ButtonUpload['state'] = "active"
                        This.Force['state'] = "disabled"
                    else:
                        This.ButtonUpload['state'] = "active"
                        This.var1.set(1)
                        This.Force['state'] = "disabled"
                    This.UpdateGuide("Etape 4 : Importez dans Alfresco")
                    return True
                else:
                    This.Force['state'] = "disabled"
                    This.var1.set(0)
                    This.logger("\nTest des destinations : dossiers manquants (sans conséquences en mode Bulk Import Tool)","")
                    This.logger("\nPour le mode CMIS (mise à jour), le mode Bulk Import Tool doit être lancé en premier)","")
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
                wkspace = unicode(row[This.confpack['wkspace']], "iso8859_1")
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

                This.fileinfo[idx]['wkspace'] = wkspace
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
    
    def initiateBulkImport(This):
        This.logger("\nTransfert du package sur "+This.conf['host']+" (Bulk Import Tool). Patientez...\n","")
        if ( This.transfertPack() ):
            This.logger("Transfert du package sur "+This.conf['host']+" OK\n","Success")
            
            rh = RESTHelper()
            rh.login(This.conf['user'], This.conf['password'], This.conf['host'], 8080)

            try:
                rh.initiateBulkImport(This.conf['importdir']+"/"+This.confpack['PKGID']+"/","/")
                This.logger("Initialisation de l'import (Bulk Import Tool)\n","Success")
                This.logger("Voir le statut : http://"+This.conf['host']+":"+This.conf['port']+"/alfresco/s/bulk/import/status\n","")
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


