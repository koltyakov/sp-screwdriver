# Screwdriver for SharePoint

> Beta | Not for production (!) | Experimental sandbox

[![NPM](https://nodei.co/npm/sp-screwdriver.png?mini=true&downloads=true&downloadRank=true&stars=true)](https://nodei.co/npm/sp-screwdriver/)

[![npm version](https://badge.fury.io/js/sp-screwdriver.svg)](https://badge.fury.io/js/sp-screwdriver)
[![Downloads](https://img.shields.io/npm/dm/sp-screwdriver.svg)](https://www.npmjs.com/package/sp-screwdriver)

![logo](https://github.com/koltyakov/sp-screwdriver/blob/master/doc/img/screwdriver-logo.png?raw=true)

> Adds missing and abstracts SharePoint APIs for transparent usage in Node.js applications

SharePoint REST API is cool, but there are cases, then it's limited or even absent (e.g. MMD is not reachable trough REST API). 

This library implements (or at least tries) some vital capabilities by wrapping legacy but still working SOAP services and by hacking HTTP requests mimicing JSOM/CSOM.

## Supported SharePoint versions:
- SharePoint Online
- SharePoint 2013
- SharePoint 2016

## APIs

### User Profiles Service

- getUserProfileByName (SOAP, /_vti_bin/UserProfileService.asmx)
- modifyUserPropertyByAccountName (SOAP, /_vti_bin/UserProfileService.asmx)
- getUserPropertyByAccountName (SOAP, /_vti_bin/UserProfileService.asmx)
- getUserProfilePropertyFor (REST, /_api/sp.userprofiles.peoplemanager/getpropertiesfor)
- getPropertiesFor (REST, /_api/sp.userprofiles.peoplemanager/getuserprofilepropertyfor)
- setSingleValueProfileProperty (HTTP, /_vti_bin/client.svc/ProcessQuery)
- setMultiValuedProfileProperty (HTTP, /_vti_bin/client.svc/ProcessQuery)

### Manage Metadata Service (Taxonomy)

- getTermSets (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getChildTermsInTermSet (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getChildTermsInTerm (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getTermsByLabel (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getKeywordTermsByGuids (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- addTerms (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getAllTerms (HTTP, /_vti_bin/client.svc/ProcessQuery)
- setTermName (HTTP, /_vti_bin/client.svc/ProcessQuery)
- deprecateTerm (HTTP, /_vti_bin/client.svc/ProcessQuery)

### Possible SOAP services to implement

- Alerts (/_vti_bin/alerts.asmx)
- Authentication Web service (/_vti_bin/Authentication.asmx)
- BDC Web Service (/_vti_bin/businessdatacatalog.asmx)
- CMS Content Area Toolbox Info Web service (/_vti_bin/contentAreaToolboxService.asmx)
- Copy Web service (/_vti_bin/Copy.asmx)
- Document Workspace Web service (/_vti_bin/DWS.asmx)
- Excel Services Web service (/_vti_bin/ExcelService.asmx)
- Meetings Web service (/_vti_bin/Meetings.asmx)
- People Web service (/_vti_bin/People.asmx)
- Permissions Web service (/_vti_bin/Permissions.asmx)
- Published Links Web service (/_vti_bin/publishedlinksservice.asmx)
- Publishing Service Web service (/_vti_bin/PublishingService.asmx)
- Search Web service (/_vti_bin/search.asmx)
- SharePoint Directory Management Web service (/_vti_bin/sharepointemailws.asmx)
- Sites Web service (/_vti_bin/sites.asmx)
- Search Crawl Web service (/_vti_bin/spscrawl.asmx)
- Users and Groups Web service (/_vti_bin/UserGroup.asmx)
- User Profile Change Web service (/_vti_bin/userprofilechangeservice.asmx)
- User Profile Web service (/_vti_bin/userprofileservice.asmx)
- Versions Web service (/_vti_bin/versions.asmx)
- Views Web service (/_vti_bin/Views.asmx)
- Web Part Pages Web service (/_vti_bin/webpartpages.asmx)
- Webs Web service (/_vti_bin/Webs.asmx)
- Workflow Web service (/_vti_bin/workflow.asmx)

[...](https://msdn.microsoft.com/en-us/library/office/bb862916(v=office.12).aspx)