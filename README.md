# Import CMIS Alfresco - Bulk Import in Alfresco via CMIS
Cette application simple permet d’effectuer la gestion des imports en masse de documents PDF dans un entrepôt Alfresco.
Elle s'appuie sur CMIS et les API Rest d'Alfresco. Le package est compatible avec le projet [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import).

Les documents PDF sont préparés sous la forme d'un package qui décrit les métadonnées qui les accompagneront.

## Origine de ce projet

Cette application n'entre pas en concurrence avec [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import). Elle vient en complément.

A l'origine, je cherchais une solution d'import en masse de documents. Le cahier des charges était le suivant :

- Possibilité d'importer une grande quantité de documents
- Les documents devaient être accompagnés des métadonnées (aspects, priopriétés et leur valeurs, tags, etc.)
- Les métadonnées devaient être renseignées par les services métiers

S'agissant de l'import de document, l'application était initialement prévue pour utiliser CMIS et REST. Toutefois, j'ai pris connaissance du projet [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import).

[Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import) est une solution puissante et très performante d'import en masse dans Alfresco. Elle prend la forme d'un module AMP à déployer sur l'instance Alfresco (WAR).

Cependant, bien que [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import) propose l'importation avec métadonnées, la mise en oeuvre n'est pas simple dans la mesure ou les métadonnées de chaque document sont stockées dans un fichier XML.

En effet, la génération de ces fichiers XML ne peut pas être confiée aux services métiers.

Dés le départ, j'ai souhaité que les services prépare l'importation en décrivant les documents et leurs métadonnées dans un tableau type Excel ou Open Office. Le fait qu'un utilisateur métier travaille avec un tableur est une solution plus "naturelle" et confortable.

Cette application s'appuie sur ce tableau pour générer un package d'importation compatible avec [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import).

Un mode CMIS est disponible permettant certaines opérations non proposées par [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import).

## Que fait cette application ?

- Elle préparer les packages d'importation
- Elle génère les packages pour l'importation via [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import)
- Elle transfert les package sur le serveur Alfresco dans un dossier défini
- Elle déclenche l'importation via [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import) (via REST)
- Elle récupère le statut de l'importation des packages (via REST)

**Licence :**
L'application est sous licence [LGPL v3.0](http://www.gnu.org/licenses/lgpl-3.0.html). 

**Version :**
Version actuelle : 1.0.1

**Compatibilité :**
La version actuelle fonctionne avec Alfresco 5.0 et 5.1. (Je n'ai pas testé avec les versions antérieures).

**Langage :**
Cette application est développée en Python 2.7.

Elle fonctionne sous Linux et Windows.

# Préparations

## Les prérequis
L'application nécéssite Python 2.7.

Les modules suivants doivent être installés :

- Tk
- Tix
- iso8601

En cas d'oublie, n'hésitez pas à m'en faire part.

L'application utilise un __aspect__ "Import_Aflresco" (ialf:package) qui doit être déployé. Le modèle incluant l'aspect est fournit (fichier Model/Import_Alfresco.ZIP) et peut être importé via le gestionnaire de modèle de la console __Outils d'Administration__ de Share.

Vous devez avoir déployé [Bulk Import Tool](https://github.com/pmonks/alfresco-bulk-import).

## Configurer l'application

La configuration globale de l'application s'effectue depuis l'interface GUI dans __Général/Configuration__.

TODO

## Préparer le package

Un package permet d'importer des documents PDF de même nature. Autrement dit sur lesquels on appliquera les mêmes aspects et renseignera les mêmes propriétés (exemple : un lot de factures, de notes de service, etc.).

Un package prend la forme d'un dossier contenant deux sous-dossiers :

- Orig
- Conf

Le dossier __Orig__ contient les documents PDF.

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

TODO

#### Le fichier __package.conf__

Il commence toujours par la balise [PACKAGE]

Ce fichier contient 5 valeurs obligatoires :

- OLDPDFNAME : le numéro de colonne contenant la liste de nom d'origine des PDF (contenu de __Orig__)
- NEWPDFNAME : le numéro de colonne contenant les nom de PDF tels qu'ils seront importés dans Alfresco (contenu généré de __PDF__)
- TITLE      : le numéro de colonne contenant le titre
- DESC       : le numéro de colonne contenant la description
- TAGS       : le numéro de colonne contenant les tags séparés par des virgules
- DESTPATH   : le numéro de colonne contenant le chemin du dossier de destination (pour compatibilité Bulk Import Tool). Le dossier na pas besoin d'exister dans le repo Alfresco. Il sera automatiquement créé par Bulk Import Tool. Exemple : /Sites/mysite/documentLibrary

Exemple :

    [PACKAGE]
    OLDPDFNAME=1
    NEWPDFNAME=2
    TITLE=3
    DESC=4
    TAGS=5
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

TODO
