# Screwdriver for SharePoint

[![NPM](https://nodei.co/npm/sp-screwdriver.png?mini=true&downloads=true&downloadRank=true&stars=true)](https://nodei.co/npm/sp-screwdriver/)

[![npm version](https://badge.fury.io/js/sp-screwdriver.svg)](https://badge.fury.io/js/sp-screwdriver)
[![Downloads](https://img.shields.io/npm/dm/sp-screwdriver.svg)](https://www.npmjs.com/package/sp-screwdriver)
[![Gitter chat](https://badges.gitter.im/gitterHQ/gitter.png)](https://gitter.im/sharepoint-node/Lobby)

![logo](https://github.com/koltyakov/sp-screwdriver/blob/master/doc/img/screwdriver-logo.png?raw=true)

> Adds missing and abstracts SharePoint APIs for transparent usage in Node.js applications

SharePoint REST API is cool, but there are cases, then it's limited or even absent (e.g. MMD is not reachable trough REST API). 

This library implements (or at least tries) some vital capabilities by wrapping legacy but still working SOAP services and by hacking HTTP requests mimicing JSOM/CSOM.

## New in version 1.0.0

- Code base is completely migrated to TypeScript.
- node-sp-auth-config is integrated to the library.
- Integration tests are added.

## Supported SharePoint versions

- SharePoint Online
- SharePoint 2013
- SharePoint 2016

## Usage

### Install

```bash
npm install sp-screwdriver --save
```

or

```bash
yarn add sp-screwdriver
```

### Minimal setup

```javascript
import { Screwdriver, IScrewdriverSettings } from 'sp-screwdriver';

const settings: IScrewdriverSettings = {
  // ...
};

const screw = new Screwdriver(settings);

// Wizard mode asks for credentials
screw.wizard().then(() => {

  screw.ups.getPropertiesFor({
    accountName: 'i:0#.f|membership|username'
  }).then(result => {
    // ...
  }).catch(console.log);

  screw.mmd.getAllTerms({
    serviceName: 'Taxonomy_5KSgChEZ9j15+7UVInQNRQ==',
    termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f'
  }).then(result => {
    // ...
  }).catch(console.log);

});
```

alternative:

```javascript
import { Screwdriver } from 'sp-screwdriver';

const screw = new Screwdriver(); // Default settings
screw.init(); // private.json already should be on the disk
              // or raw auth parameters should be provided

screw.ups.getUserPropertyByAccountName({
  accountName: 'i:0#.f|membership|username',
  propertyName: 'SPS-Birthday'
}).then(result => {
  done();
}).catch(done);
```

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

### Versions

#### Document versions

- getVersions (SOAP, /_vti_bin/versions.asmx)
- restoreVersion (SOAP, /_vti_bin/versions.asmx)
- deleteVersion (SOAP, /_vti_bin/versions.asmx)
- deleteAllVersions (SOAP, /_vti_bin/versions.asmx)

#### Item versions

- getVersionCollection (SOAP, /_vti_bin/lists.asmx)

#### Item property bags

- setItemProperties (HTTP, /_vti_bin/client.svc/ProcessQuery)

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
- Views Web service (/_vti_bin/Views.asmx)
- Web Part Pages Web service (/_vti_bin/webpartpages.asmx)
- Webs Web service (/_vti_bin/Webs.asmx)
- Workflow Web service (/_vti_bin/workflow.asmx)

[...](https://msdn.microsoft.com/en-us/library/office/bb862916(v=office.12).aspx)
