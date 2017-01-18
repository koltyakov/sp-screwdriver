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

### Manage Metadata Service (Taxonomy)

- getUserProfileByName (SOAP, /_vti_bin/UserProfileService.asmx)
- modifyUserPropertyByAccountName (SOAP, /_vti_bin/UserProfileService.asmx)
- getUserPropertyByAccountName (SOAP, /_vti_bin/UserProfileService.asmx)
- getUserProfilePropertyFor (REST, /_api/sp.userprofiles.peoplemanager/getpropertiesfor)
- getPropertiesFor (REST, /_api/sp.userprofiles.peoplemanager/getuserprofilepropertyfor)

### User Profiles Service

- getTermSets (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getChildTermsInTermSet (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getChildTermsInTerm (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getTermsByLabel (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getKeywordTermsByGuids (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- addTerms (SOAP, /_vti_bin/TaxonomyClientService.asmx)
- getAllTerms (HTTP, /_vti_bin/client.svc/ProcessQuery)
- setTermName (HTTP, /_vti_bin/client.svc/ProcessQuery)
- deprecateTerm (HTTP, /_vti_bin/client.svc/ProcessQuery)