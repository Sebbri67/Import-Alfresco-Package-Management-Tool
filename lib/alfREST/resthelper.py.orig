# encoding: utf-8
# Copyright (c) 2011 AXIA Studio <info@axiastudio.it>
#
# This file may be used under the terms of the GNU General Public
# License versions 3.0 or later as published by the Free Software
# Foundation and appearing in the file LICENSE.
#
# This file is provided AS IS with NO WARRANTY OF ANY KIND, INCLUDING THE
# WARRANTY OF DESIGN, MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.
#

import httplib
import socket
from datetime import datetime
from xml.dom.minidom import parseString, getDOMImplementation
import logging

from webscripts import *
from formdata import get_content_type, encode_multipart_formdata


ROLES = dict(
CONSUMER = "{http://www.alfresco.org/model/content/1.0}cmobject.Consumer",
EDITOR = "{http://www.alfresco.org/model/content/1.0}cmobject.Editor",
CONTRIBUTOR = "{http://www.alfresco.org/model/content/1.0}cmobject.Contributor",
COLLABORATOR = "{http://www.alfresco.org/model/content/1.0}cmobject.Collaborator",
COORDINATOR = "{http://www.alfresco.org/model/content/1.0}cmobject.Coordinator",
)

CMIS_PROPERTY_TYPES = {"cmis:propertyString"    : (lambda t: t),
                       "cmis:propertyDateTime"  : (lambda t: datetime.strptime(t[:19], "%Y-%m-%dT%H:%M:%S")),
                       "cmis:propertyId"        : (lambda t: t),
                       "cmis:propertyBoolean"   : bool,
                       "cmis:propertyInteger"   : int,
                       }

XMLNS_CMIS = "http://docs.oasis-open.org/ns/cmis/core/200908/"
XMLNS_CMISRA = "http://docs.oasis-open.org/ns/cmis/restatom/200908/"

logger = logging.getLogger("alfrescohelper")
logger.setLevel(logging.DEBUG)
sh = logging.StreamHandler()
sh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(sh)


def parseComm(s):
    method, comm = s.split(" ")
    comm = comm.replace("{", "%(")
    comm = comm.replace("?}", ")s")
    comm = comm.replace("}", ")s")
    return method, comm


def formatComm(comm, pars):
    try:
        comm.index("?")
    except ValueError:
        return comm % pars
    noneKeys = [k for k in pars.keys() if pars[k] is None]
    for k in noneKeys:
        del pars[k]
    if len(pars)==0:
        return comm
    url, urlPars = comm.split("?")
    tkns = urlPars.split("&")
    outTkns = []
    for tkn in tkns:
        try:
            outTkns.append(tkn % pars)
        except KeyError:
            pass
    newComm = url + "?" + "&".join(outTkns)
    return newComm % pars


def pythonizeJson(json):
    json = json.replace("\r", "")
    json = json.replace("true", "True")
    json = json.replace("false", "False")
    json = json.replace("null", "None")
    return eval(json, {}, {})


def extractCmisProperties(node):
    child = {}
    for prop in node.getElementsByTagName("cmis:properties")[0].childNodes:
        try:
            attr = prop.getAttribute("propertyDefinitionId")
        except AttributeError:
            attr = None
        if attr is not None:
            values = prop.getElementsByTagName("cmis:value")
            if len(values) == 1:
                try:
                    value = CMIS_PROPERTY_TYPES[prop.tagName](values[0].childNodes[0].wholeText)
                    child[attr] = value
                except KeyError:
                    pass
    return child

def restfuldoc(method):
    """
    Decore method's docstring with RESTful help
    """
    method.__doc__ = "\n"+globals()[method.__name__.upper()][0]
    return method


class RESTHelper(object):

    def __init__(self):
        self.ticket = None
        self.host = "127.0.0.1"
        self.port = "8080"

    @restfuldoc
    def login(self, username, password, host="127.0.0.1", port="8080"):

        self.host = host
        self.port = port
        jsonDict = dict(username=username, password=password)
        method, url = parseComm(LOGIN[1])
        res = self.request(method, url, json=repr(jsonDict))
        if res is not None:
            data = res.read()
            json = pythonizeJson(data)
            self.ticket = json["data"]["ticket"]
            logger.info("Alfresco login.")
            return True
        return False

    @restfuldoc
    def logout(self):

        getPars = dict(ticket=self.ticket)
        method, url = parseComm(LOGOUT[1])
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info("Alfresco logout.")
            return True
        return False

    @restfuldoc
    def listAllRootGroups(self, shortNameFilter=None, zone=None, maxItems=None, skipCount=None,
                                sortBy=None):

        method, url = parseComm(LISTALLROOTGROUPS[1])
        getPars = dict(shortNameFilter = shortNameFilter,
                       zone = zone,
                       maxItems = maxItems,
                       skipCount = skipCount,
                       sortBy = sortBy)
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info("List all root groups.")
            groups = pythonizeJson(res.read())["data"]
            return groups
        return False


    @restfuldoc
    def getACL(self, path):

        getPars = dict(path=path)
        method, url = parseComm(GETACL[1])
        res = self.request(method, url, getPars=getPars)
        acl = {}
        if res is not None:
            logger.info("Get ACL for the path %s." % path)
            doc = parseString(res.read())
            #print doc.toxml()
            for child in doc.documentElement.childNodes:
                if child.nodeName == "cmis:permission":
                    principal = child.getElementsByTagName("cmis:principal")[0]
                    principiaId = principal.getElementsByTagName("cmis:principalId")[0].childNodes[0].wholeText
                    permissionNodes = child.getElementsByTagName("cmis:permission")
                    permissions = []
                    for permissionNode in permissionNodes:
                        permissions.append(permissionNode.childNodes[0].wholeText)
                    direct = child.getElementsByTagName("cmis:direct")[0].childNodes[0].wholeText
                    if principiaId not in acl.keys():
                        acl[principiaId] = []
                    acl[principiaId].append((permissions, direct=="true"))
        return acl


    @restfuldoc
    def applyACL(self, path, acl):

        impl = getDOMImplementation()
        doc = impl.createDocument(XMLNS_CMIS, "cmis:acl", None)
        doc.documentElement.setAttribute("xmlns:cmis", XMLNS_CMIS)
        for principiaIdText in acl.keys():
            for permissions, direct in acl[principiaIdText]:
                permissionElement = doc.createElement("cmis:permission")
                doc.documentElement.appendChild(permissionElement)
                principalElement = doc.createElement("cmis:principal")
                permissionElement.appendChild(principalElement)
                principalIdElement = doc.createElement("cmis:principalId")
                principalElement.appendChild(principalIdElement)
                principalIdElement.appendChild(doc.createTextNode(principiaIdText))
                for permission in permissions:
                    permissionSubElement = doc.createElement("cmis:permission")
                    permissionSubElement.appendChild(doc.createTextNode(permission))
                    permissionElement.appendChild(permissionSubElement)
                    directElement = doc.createElement("cmis:direct")
                    permissionElement.appendChild(directElement)
                    directElement.appendChild(doc.createTextNode(direct is True and 'true' or 'false'))

        getPars = dict(path=path)
        method, url = parseComm(APPLYACL[1])

        res = self.request(method, url, getPars=getPars, cmisaclxml=doc.toxml())
        if res is not None:
            logger.info("ACL applied to %s." % path)


    @restfuldoc
    def listChildAuthorities(self, shortName, authorityType='USER'):

        getPars = dict(shortName=shortName, authorityType=authorityType)
        method, url = parseComm(LISTCHILDAUTHORITIES[1])
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info('List child authorities (%s) for the group %s.' % (authorityType, shortName))
            users = pythonizeJson(res.read())['data']
            return users
        return False


    @restfuldoc
    def addPerson(self, userName, firstName, lastName, email, password=None):

        pars = dict(userName=userName, firstName=firstName, lastName=lastName, email=email)
        if password is not None:
            pars['password'] = password
        method, url = parseComm(ADDPERSON[1])
        res = self.request(method, url, json=repr(pars))
        if res is not None:
            logger.info('Person %s (%s %s) added.' % (userName, firstName, lastName))
            return True
        return False

    @restfuldoc
    def deletePerson(self, userName):

        getPars = dict(userName=userName)
        method, url = parseComm(DELETEPERSON[1])
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info('Person %s deleted.' % userName)
            return True
        return False


    @restfuldoc
    def addRootGroup(self, shortName, displayName=None):
        getPars = dict(shortName=shortName,)
        if displayName is not None:
            jsonDict = dict(displayName=displayName)
        else:
            jsonDict = {}
        method, url = parseComm(ADDROOTGROUP[1])
        res = self.request(method, url, getPars=getPars, json=repr(jsonDict))
        if res is not None:
            logger.info('Root group %s (%s) added' % (shortName, displayName))
            return True
        return False


    @restfuldoc
    def deleteRootGroup(self, shortName):

        getPars = dict(shortName=shortName,)
        method, url = parseComm(DELETEROOTGROUP[1])
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info('Root group %s deleted.' % shortName)
            return True
        return False


    @restfuldoc
    def addGroupOrUserToGroup(self, fullAuthorityName, shortName):

        getPars = dict(fullAuthorityName=fullAuthorityName, shortName=shortName)
        method, url = parseComm(ADDGROUPORUSERTOGROUP[1])
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info('Group or user %s inserted in group %s.' % (fullAuthorityName, shortName))
            return True
        return False


    @restfuldoc
    def removeAuthorityFromGroup(self, fullAuthorityName, shortGroupName):

        getPars = dict(fullAuthorityName=fullAuthorityName, shortGroupName=shortGroupName)
        method, url = parseComm(REMOVEAUTHORITYFROMGROUP[1])
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info('Authority %s removed from group %s.' % (fullAuthorityName, shortGroupName))
            return True
        return False


    @restfuldoc
    def fileUpload(self, filedata, siteid, containerid, uploaddirectory, updatenoderef=None,
            description=None, contenttype=None, majorversion=None, overwrite=None, thumbnails=None):

        filename = filedata.name
        contenttype = get_content_type(filename)
        pars = dict(siteid = siteid,
                    containerid = containerid,
                    uploaddirectory = uploaddirectory,
                    filename = filename,
                    contenttype = contenttype)
        files = (("filedata", filename, filedata.read()),)
        method, url = parseComm(FILEUPLOAD[1])

        content_type, body = encode_multipart_formdata(pars.items(), files)

        headers = {
            "User-Agent": "alfREST",
            "Host": "%s:%s" % (self.host, self.port),
            "Accept": "*/*",
            "Content-Length": str(len(body)),
            "Expect": "100-continue",
            "Content-Type": content_type,
        }

        res = self.request(method, url, body, headers)
        if res is not None:
            logger.info('File %s uploaded in %s/%s%s.' % (filename, siteid, containerid, uploaddirectory))
            return eval(res.read())["nodeRef"].split("/")[-1]
        return False


    @restfuldoc
    def createFolder(self, path, name, summary=""):

        impl = getDOMImplementation()
        doc = impl.createDocument(XMLNS_CMIS, "entry", None)
        doc.documentElement.setAttribute("xmlns", "http://www.w3.org/2005/Atom")
        doc.documentElement.setAttribute("xmlns:cmis", XMLNS_CMIS)
        doc.documentElement.setAttribute("xmlns:cmisra", XMLNS_CMISRA)
        title = doc.createElement("title")
        title.appendChild(doc.createTextNode(name))
        doc.documentElement.appendChild(title)
        idNode = doc.createElement("summary")
        idNode.appendChild(doc.createTextNode("ignored"))
        doc.documentElement.appendChild(idNode)
        summaryNode = doc.createElement("summary")
        summaryNode.appendChild(doc.createTextNode(summary))
        doc.documentElement.appendChild(summaryNode)
        object = doc.createElement("cmisra:object")
        doc.documentElement.appendChild(object)
        properties = doc.createElement("cmis:properties")
        object.appendChild(properties)

        propertyId = doc.createElement("cmis:propertyId")
        propertyId.setAttribute("propertyDefinitionId", "cmis:objectTypeId")
        properties.appendChild(propertyId)
        value = doc.createElement("cmis:value")
        value.appendChild(doc.createTextNode("cmis:folder"))
        propertyId.appendChild(value)

        propertyString = doc.createElement("cmis:propertyString")
        propertyString.setAttribute("propertyDefinitionId", "cmis:name")
        properties.appendChild(propertyString)
        nameValue = doc.createElement("cmis:value")
        nameValue.appendChild(doc.createTextNode(name))
        propertyString.appendChild(nameValue)

        method, url = parseComm(CREATEFOLDER[1])
        body = doc.toxml()
        headers = {
            "User-Agent": "alfREST",
            "Host": "%s:%s" % (self.host, self.port),
            "Accept": "*/*",
            "Content-Length": str(len(body)),
            "Content-Type": "application/atom+xml",
        }

        pars = {"path":path,}
        res = self.request(method, url, body, headers, pars)
        if res is not None:
            logger.info("Folder %s created in %s." % (name, path))
            doc = parseString(res.read())
            propertyIds = doc.getElementsByTagName("cmis:propertyId")
            for pid in propertyIds:
                try:
                    attr = pid.getAttribute("propertyDefinitionId")
                    if attr == "cmis:objectId":
                        objectId = pid.childNodes[0].childNodes[0].wholeText.split("/")[-1]
                        return objectId
                except AttributeError:
                    pass


    @restfuldoc
    def deleteObject(self, idObject):

        getPars = dict(id=idObject, property="")
        method, url = parseComm(DELETEOBJECT[1])
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info('Object %s removed.' % idObject)
            return True
        return False


    @restfuldoc
    def deleteContent(self, idObject):

        getPars = dict(id=idObject, property="")
        method, url = parseComm(DELETECONTENT[1])
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info('Content %s removed.' % idObject)
            return True
        return False


    @restfuldoc
    def getChildren(self, path):

        method, url = parseComm(GETCHILDREN[1])
        getPars = dict(path = path)
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info("Gets the list of child objects contained in %s." % path)
            doc = parseString(res.read())
            children = []
            for childNode in doc.getElementsByTagName("entry"):
                child = extractCmisProperties(childNode)
                children.append(child)
            return children
        return False


    @restfuldoc
    def getProperties(self, path):

        method, url = parseComm(GETPROPERTIES[1])
        getPars = dict(path = path)
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info("Gets the list of child objects contained in %s." % path)
            doc = parseString(res.read())
            objectNode = doc.getElementsByTagName("cmisra:object")[0]
            prop = extractCmisProperties(objectNode)
            return prop
        return False


    @restfuldoc
    def getContent(self, pathOrId, property=False):

        if pathOrId[0] == "/":
            method, url = parseComm(GETCONTENT[1])
            getPars = dict(path = pathOrId, property = "property")
        else:
            method, url = parseComm(GETCONTENT[2])
            getPars = dict(id = pathOrId, property = "property")
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info("Gets content from %s." % pathOrId)
            return res.read()
        return False


    @restfuldoc
    def getRendition(self, path, property=False):

        return self.getContent(path, True)


    @restfuldoc
    def getNodeTags(self, store_type, store_id, id_):

        method, url = parseComm(GETNODETAGS[1])
        getPars = dict(store_type=store_type, store_id=store_id, id=id_)
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info("Get all the tags for node %s://%s/%s." % (store_type, store_id, id_))
            content = res.read()
            if content.strip() == 'False':
                return False
            tags = []
            for row in content.split('\n'):
                row = row.strip()
                if row not in ('[', ']'):
                    tags.append(row)
            return tags
        return False


    @restfuldoc
    def addTag(self, store_type, store_id, id_, tag):

        method, url = parseComm(ADDTAG[1])
        getPars = dict(store_type=store_type, store_id=store_id, id=id_)
        res = self.request(method, url, getPars=getPars, json='["%s"]'%tag)
        if res is not None:
            logger.info('Tag %s added to %s.' % (tag, id_))
            return True
        return False


    @restfuldoc
    def deleteTag(self, store_type, store_id, id_, tag):

        method, url = parseComm(DELETETAG[1])
        getPars = dict(store_type=store_type, store_id=store_id, id=id_, tagName=tag)
        res = self.request(method, url, getPars=getPars)
        if res is not None:
            logger.info('Tag %s deleted form %s.' % (tag, id_))
            return True
        return False


    def request(self, method, url, body=None, headers=None, getPars={}, json=None, cmisaclxml=None, atomxml=None):

        # auth ticket in url
        if self.ticket is not None:
            try:
                url.index("?")
                c = "&"
            except ValueError:
                c = "?"
            url += "%salf_ticket=%s" % (c, self.ticket)

        # body and headers for json and cmisaclxml
        if json is not None:
            headers = {"Content-Type":"application/json", "Accept":"text/plain"}
            body = json
        elif cmisaclxml is not None:
            headers = {"Content-Type":"application/cmisacl+xml", "Accept":"text/plain"}
            body = cmisaclxml
        elif atomxml is not None:
            headers = {"Content-Type":"application/atom+xml;type=entry", "Accept":"text/plain"}
            body = atomxml

        http_conn = httplib.HTTPConnection(self.host, self.port, timeout=60)

        if method in ("GET", "DELETE"):
            http_conn.request(method, formatComm(url, getPars))

        elif method in ("POST", "PUT"):
            if headers is None:
                if body is None:
                    http_conn.request(method, formatComm(url, getPars))
                else:
                    http_conn.request(method, formatComm(url, getPars), body)
            else:
                http_conn.request(method, formatComm(url, getPars), body, headers)

        try:
            res = http_conn.getresponse()
        except socket.timeout:
            logger.error("Server timeout")
            return None
        if res.status/100 in (4, 5):
            logger.error(res.reason)
            content = res.read()
            if content[:9] == "<!DOCTYPE":
                doc = parseString(content)
                title = doc.getElementsByTagName("title")[0]
                msg = title.firstChild.wholeText
            else:
                try:
                    msg = eval(content)["message"]
                except SyntaxError:
                    msg = "Unable to catch error message"
            logger.error(msg)
            return None
        return res

