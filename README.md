# Import CMIS Alfresco - Bulk Import in Alfresco via CMIS
Cette application simple permet d’effectuer des imports en masse de documents PDF dans un entrepôt Alfresco.
Elle s'appuie sur CMIS et les API Rest d'Alfresco. Le package est aussi compatible avec le projet [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import).

Les documents PDF sont préparés sous la forme d'un package qui décrit les métadonnées qui les accompagneront.

Import Alfresco permet en effet d'importer les PDF et de leur affecter des Aspects, des priopriétés ainsi que des tags éventuels.

L'importation s'effectue en 4 étapes :

- Préparation du package
- Préparation des PDF (renommage en masse et inclusion des métadonnées Title et Description)
- Choix du serveur Alfresco
- Importation des PDF dans Alfresco via CMIS __OU__ le projet [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import)

__A noter__ : Le projet Bulk Import Tool (https://github.com/pmonks/alfresco-bulk-import) est beaucoup plus performant car il n'utilise pas CMIS. Par ailleurs, il ne se limite pas aux PDF.

En revanche, avec ce dernier, les métadonnées doivent être décrites dans un fichier XML accompagnant chaque document.

La création de ces fichiers XML peut être une tâche fastidieux. Aussi, l'action "Générer les PDFs" génère les fichiers XML afin que le package puisse être compatible avec Bulk Import Tool.

En cas d'une grosse quantité de documents à importer, cette autre solution sera plus pertinante.

## Pourquoi cette nouvelle application ?

Cette application n'entre pas en concurrence avec Bulk Import Tool. Cette dernière solution est particulièrement efficace mais exige la création des XML Metadata pour chaque document importé. Ceci peut être fastidieux.

Pour l'import avec métadonnées (aspects, properties, tags, etc.), il me semblait plus aisé de préparer l'import depuis un simple tableau (CSV) plutôt que de devoir écrire chaque XML Metadata.

Autrement dit, mon application vient en complément.

![Import Alfresco](https://github.com/Sebbri67/Import-Alfresco/blob/master/Import_Alfresco.jpg)

**Licence :**
L'application est sous licence [LGPL v3.0](http://www.gnu.org/licenses/lgpl-3.0.html). 

**Version :**
Version actuelle : 1.0.1

**Compatibilité :**
La version actuelle fonctionne avec Alfresco 5.0 et 5.1. (Je n'ai pas testé avec les versions antérieures).

**Langage :**
Cette application est développée en Python 2.7.

Elle fonctionne sous Linux et Windows.

Elle s'appuie sur trois librairies (incluses) :

- Cmislib et une extention Cmislibalf
- alfREST
- PyPDF2

# Préparations

## Les prérequis
L'application nécéssite Python 2.7.

Les modules suivants doivent être installés :

- Tk
- Tix
- iso8601

Dans Alfresco, un dossier __Imports__ doit être créé à la racine de l'entrepôt (au même niveau que Sites).

L'application utilise un __aspect__ "Import_Aflresco" (ialf:package) qui doit être déployé. Le modèle incluant l'aspect est fournit (fichier Import_Alfresco.ZIP) et peut être importé via le gestionnaire de modèle de la console __Outils d'Administration__ de Share.

## Configurer l'application

La configuration globale de l'application s'effectue dans le fichier __importalf.conf__. Il peut être modifié depuis l'interface GUI dans __Général/Configuration__.

Il commence toujours par la balise [GLOBAL]

Quatres variables doivent être renseignées :

- L'url Template ( __urltemp__ ) qui correspond à l'url d'accès de l'API CMIS d'Alfresco. Il faut remplacer le nom d'hôte par ____HOST____
- Le login utilisateur ( __user__ ) pour l'authentification (exemple : admin).
- Le mot de passe ( __password__ ).
- La liste de vos hôtes Alfresco ( __hosts__ ) (exemple  : ged.yourdomain.net). Les hôtes doivent être séparés par des virgules.

## Préparer le package

Un package permet d'importer des documents PDF de même nature. Autrement dit sur lesquels on appliquera les mêmes aspects et renseignera les mêmes propriétés (exemple : un lot de factures, de notes de service, etc.).

Un package prend la forme d'un dossier contenant trois sous-dossiers :

- Orig
- PDF
- Conf

Le dossier __Orig__ contient les documents PDF.

Le dossier __PDF__ contiendra les documents PDF après traitement (renommage et métadonnées mises à jour). Ce sont ces documents qui seront véritablement importés dans Alfresco.

Le dossier __Conf__ contient la configuration du package.

Un exemple de package est fournit.

###Configuration du package

Le dossier __Conf__ doit contenir obligatoirement 4 fichiers.

- Le fichier __list.csv__
- Le fichier __package.conf__
- Le fichier __aspects.conf__
- Le fichier __properties.csv__

Un cinquième fichier nommé __packageId__ sera généré lors de la première ouverture. Il contient le Package ID qui ne devra jamais être modifier ni supprimé si vous souhaitez relancer le traitement plusieurs fois (ajouts de PDF, modification des aspects et/ou propriétés, etc.).

####Le fichier __list.csv__ :

Il contient la liste des documents ainsi que toutes les informations nécéssaires à l'import.

Les contraintes obligatoires :

- La première ligne doit contenir le nom des colonnes
- La première colonne doit contenir un identifiant unique (exemple : un incrément de 1). Attention, vous devrez par la suite conserver la cohérence Identifiant/Document PDF. Cet identifiant, couplé avec le package ID, servira à renseigner une propriété de l'aspect ialf:package et identifiera le document une fois importé dans Alfresco.

Ensuite, le fichier CSV doit obligatoirement posséder les colonnes suivantes :

- Une colonne contenant le nom d'origine des documents PDF (contenu de __Orig__)
- Une colonne contenant le nom du document PDF tel qu'il sera importé (contenu généré de __PDF__)
- Une colonne contenant le titre du document. Le titre servira à mettre à jour la métadonnée Title du PDF et la propriété cm:title du document dans Alfresco.
- Une colonne contenant la description du document. La description servira à mettre à jour la métadonnée Description du PDF et la propriété cm:description du document dans Alfresco.
- Une colonne contenant les tags qui seront appliqués au document dans Alfresco (cette colonne est obligatoire mais peut contenir une valeur vide). Les tags sont séparés par des virgules (sans espaces entre les mots et les virgules).
- Une colonne contenant la destination de chaque document. La destination est le nodeId du dossier dans Alfresco (exemple : c2926bde-bbe1-4e97-8c07-8fec2754e595)

D'autres colonnes peuvent être ajoutées facultativement afin de renseigner les valeurs des propriétés personnalisées (voir __properties.csv__).

__A NOTER__ : les colonnes du CSV sont numérotées de 0 à X

Exemple : 

    Voir l'exemple dans le dossier "Exemple"

#### Le fichier __package.conf__

Il commence toujours par la balise [PACKAGE]

Ce fichier contient 6 valeurs obligatoires :

- OLDPDFNAME : le numéro de colonne contenant la liste de nom d'origine des PDF (contenu de __Orig__)
- NEWPDFNAME : le numéro de colonne contenant les nom de PDF tels qu'ils seront importés dans Alfresco (contenu généré de __PDF__)
- TITLE      : le numéro de colonne contenant le titre
- DESC       : le numéro de colonne contenant la description
- TAGS       : le numéro de colonne contenant les tags séparés par des virgules
- DESTWSP    : le numéro de colonne contenant le nodeId du dossier de destination (__le nodeId doit être celui d'un dossier déjà existant dans le repo Alfresco__)
- DESTPATH   : le numéro de colonne contenant le chemin du dossier de destination (pour compatibilité Bulk Import Tool). Le dossier na pas besoin d'exister dans le repo Alfresco. Il sera automatiquement créé par Bulk Import Tool. Exemple : /Sites/mysite/documentLibrary

Exemple :

    [PACKAGE]
    OLDPDFNAME=1
    NEWPDFNAME=2
    TITLE=3
    DESC=4
    TAGS=5
    DESTWSP=11
    DESTPATH=10

#### Le fichier __aspects.conf__

Il contient la liste des aspects qui seront appliqués aux documents du package.

Exemple :

    gen:classement
    mya:facture

#### Le fichier __properties.csv__

Il contient la liste de propriétés à appliquer aux documents.

- Colonne 0 : le nom de la propriété
- Colonne 1 : le format de la propriété (TXT, NUM ou DATE)
- Colonne 2 : le type de liaison (STA pour statique, DYN pour dynamique)
- Colonne 3 : la valeur que prendra la propriété. Pour une valeur statique : la valeur. Pour une valeur dynamique : le numéro de la colonne dans __list.csv__

Exemples :

    gen:statut;TXT;STA;Classé
    mya:annee;NUM;DYN;7
    mya:date;DATE;DYN;8
    mya:numfact;TXT;DYN;9
    cm:title;TXT;DYN;3
    cm:description;TXT;DYN;4

#Utilisation

L'application est sous la forme d'un client graphique (en Python Tk).

Lancez l'application par __./importalf.py__

(Sous Windows, renommez le fichier en __importalf.pyw__ afin que la console ne s'ouvre pas).

1. Configurez correctement l'application dans __Général/Configuration__
2. Etape 1 : Ouvrez votre package en cliquant sur __Choisir le package__. Vous devez sélectionner le dossier contenant les trois sous dossiez évoqués plus haut.
3. Vérifier les information, le résultat doit être __OK__. Vérifiez aussi la correspondance des champs du CSV avec les attributs du __package.conf__ et de __properties.csv__, le cas échéant.
4. Etape 2 : Cliquez sur __Générer les PDFs__. Ce processus créera une copie renommée des PDF (avec métadonnées communes à jour) dans le dossier __PDF__ ainsi que les fichiers XML Metadata.

### Import en mode CMIS :

5. Etape 3 : Choisissez l'hôte du serveur Alfresco dans la liste déroulante.
6. Vérifiez l'état du test de connexion. Le test vérifie aussi que les nodeId (colonne Destination) sont valides.
7. Etape 4 : Cliquez sur __Importer les PDFs__

Le processus d'import est lancé.

Il est possible de relancer le traitement d'un package. Par défaut, si les documents sont déjà importés, l'application ne fait rien et indique que le PDF est déjà existant dans Alfresco.

Le fait de relancer un traitement permet :

- la prise en compte de nouveaux PDF ajoutés au package
- la possibilité de mettre à jour les métadonnées des documents déjà importés. Pour cela, il faut __Forcer la mise à jour__.

La mise à jour s'applique sur les métadonnées tels que le titre, la description, les aspects (ajout), les propriétés (valeurs et ajout) et le nom du document.

Par exemple, si après avoir importé les PDF vous vous rendez compte que le nommage de ces derniers ne vous convient pas, il n'est pas utile de les supprimer dans Alfresco pour les réimporter. En effet, le nom d'un document dans Alfresco n'est qu'une simple propriété (cmis:name). L'application ne se base pas sur le nom pour retrouver un document déjà importé. Il se base sur une propriété de l'aspect ialf:package contenant un identifiant unique.

### Import en mode [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import) :

TODO

