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

LOGIN = (
"""
Login and establish a ticket.

Input
    JSON Data Object.

username
    cleartext username

password
    cleartext password

Returns the new authentication ticket.
""",
"POST /alfresco/service/api/login"
)


LOGOUT = (
"""
Logout, Delete Login Ticket.

After the user has logged out the ticket is no longer valid and subsequent attempts to use it will
fail.
""",
"DELETE /alfresco/service/api/login/ticket/{ticket}"
)


LISTALLROOTGROUPS = (
"""
List all root groups.

If the optional zone parameter is set to 'true' then returns root groups from the specified zone.
If not specified will return groups from all zones.
If the optional shortNameFilter parameter is set then returns those root groups with a partial match
on shortName. The shortname filter can contain the wild card characters * and ? but these must be
url encoded for this script. The optional maxItems parameter sets the maximum number of items to be
returned. If no value is set then all items are returned. The optional skipCount parameter
determines how many items to skip before returning the first result. If no skipCount value is set
then no items are skipped. If the optional sortBy parameter is given, then the results may be
sorted.

Possible values are "authorityName" (default), "shortName" and "displayName"
""",
"GET /alfresco/service/api/rootgroups?shortNameFilter={shortNameFilter?}&zone={zone?}&maxItems={maxItems?}&skipCount={skipCount?}&sortBy={sortBy?}"
)


GETACL = (
"""
Get the ACL currently applied to the specified document or folder object.
""",
"GET /alfresco/service/cmis/p{path}/acl"
)


APPLYACL = (
"""
Adds or removes the given ACEs to or from the ACL of document or folder object.
""",
"PUT /alfresco/service/cmis/p{path}/acl"
)


LISTCHILDAUTHORITIES = (
"""
Get a list of the child authorities of a group. The list contains both people and groups.

The authorityType parameter is used to specify return authorities of the given type. Valid values
are GROUP and USER. The optional maxItems parameter sets the maximum number of items to be returned.
If no value is set then all items are returned. The optional skipCount parameter determines how many
items to skip before returning the first result. If no skipCount value is set then no items are
skipped. If the optional sortBy parameter is given, then the results may be sorted.

Possible values are "authorityName" (default), "shortName" and "displayName"
""",
"GET /alfresco/service/api/groups/{shortName}/children?authorityType={authorityType?}&maxItems={maxItems?}&skipCount={skipCount?}&sortBy={sortBy?}"
)


ADDPERSON = (
"""
Adds a new person based on the details provided.

userName
        mandatory - the user name for the new user

firstName
        mandatory - the given Name

lastName
        mandatory - the family name

email
        mandatory - the email address

password
        optional - the new user's password. If not specified then a value of "password" is used
        which should be changed as soon as possible.

disableAccount
        optional - If present and set to "true" the user is created but their account will be
        disabled.

quota
    optional - Sets the quota size for the new user, in bytes.

groups
    optional - Array of group names to assign the new user to.

title
    optional - the title for the new user.

organisation
    optional - the organisation the new user belongs to.

jobtitle
    optional - the job title of the new user.
""",
"POST /alfresco/service/api/people"
)


DELETEPERSON = (
"""
Delete a person.
""",
"DELETE /alfresco/service/api/people/{userName}"
)


ADDROOTGROUP = (
"""
Adds a root group. 

You must have "administrator" privileges to add a root group. 
Returns STATUS_CREATED if a new group is created. 
If the group already exists returns BAD_REQUEST. The following properties may be specified for the new root group:-

displayName
        The display name
""",
"POST /alfresco/service/api/rootgroups/{shortName}"
)


DELETEROOTGROUP = (
"""
Deletes a root group and all its dependents. 
You must have "administrator" privileges to delete a group.
""",
"DELETE /alfresco/service/api/rootgroups/{shortName}"
)


ADDGROUPORUSERTOGROUP = (
"""
Adds a group or user to a group.
The webscript will create a sub group if one does not already exist, with the fullAuthorityName. 
You must have "administrator" privileges to modify groups. 
If the authority is for a group and doe not exist then it is created. 
The webscript returns Status_Created if a new group is created, otherwise it returns Status_OK.
If Status_Created returns the new sub group, otherwise returns the group.
""",
"POST /alfresco/service/api/groups/{shortName}/children/{fullAuthorityName}"
)


REMOVEAUTHORITYFROMGROUP = (
"""
Remove an authority (USER or GROUP) from a group. A user will not be deleted by this method. 
You must have "administrator" privileges to alter a group. 
""",
"DELETE /alfresco/service/api/groups/{shortGroupName}/children/{fullAuthorityName}"
)


CREATEFOLDER = (
"""
Creates a folder.
""",
"POST /alfresco/service/cmis/p{path}/children?sourceFolderId={sourceFolderId}&versioningState={versioningState?}"
)


FILEUPLOAD = (
"""
Upload file content and meta-data into repository. 

HTML form data

    filedata, (mandatory) HTML type file
    siteid
    containerid
    uploaddirectory
    updatenoderef
    description
    contenttype
    majorversion
    overwrite
    thumbnails

Return content

    nodeRef

Return status: STATUS_OK (200)
""",
"POST /alfresco/service/api/upload"
)


DELETEOBJECT = (
"""
Delete the specified object.
""",
"DELETE /alfresco/service/cmis/i/{id}"
)

DELETECONTENT = (
"""
Deletes the content stream for the specified Document object.
""",
"DELETE /alfresco/service/cmis/i/{id}/content{property}"
)

GETCHILDREN = (
"""
Gets the list of child objects contained in the specified folder.
""",
"GET /alfresco/service/cmis/p{path}/children?types={types?}&filter={filter?}&skipCount={skipCount?}&maxItems={maxItems?}&includeAllowableActions={includeAllowableActions?}&includeRelationships={includeRelationships?}&renditionFilter={renditionFilter?}&orderBy={orderBy?}"
)

GETPROPERTIES = (
"""
Gets the properties for the object (Folder or Document).
""",
"GET /alfresco/service/cmis/p{path}?filter={filter?}&returnVersion={returnVersion?}&includeAllowableActions={includeAllowableActions?}&includeRelationships={includeRelationships?}&includeACL={includeACL?}&renditionFilter={renditionFilter?}"
)

GETCONTENT = (
"""
Gets the content stream for the specified document, or gets a rendition stream for a specified rendition of a document.
""",
"GET /alfresco/service/cmis/p{path}/content{property}?a={attach?}&streamId={streamId?}",
"GET /alfresco/service/cmis/i/{id}/content{property}?a={attach?}&streamId={streamId?}"
)

GETRENDITION = GETCONTENT

GETNODETAGS = (
"""
Get all the tags for a node.
Returns STATUS_OK (200)
""",
"GET /alfresco/service/api/node/{store_type}/{store_id}/{id}/tags"
)

ADDTAG = (
"""
Add one or more tags to the node.
Input:
(mandatory) array of String
Returns the array of tags
Return STATUS_OK (200).
""",
"POST /alfresco/service/api/node/{store_type}/{store_id}/{id}/tags"
)

DELETETAG = (
"""
Delete the given tag
""",
"DELETE /alfresco/service/api/tags/{store_type}/{store_id}/{tagName}"
)