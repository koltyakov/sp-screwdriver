import { ISPRequest } from 'sp-request';
import { Utils } from './../utils';

import { IDocumentVersions, IDocumentVersion, IItemVersions } from './../interfaces/IVersions';

export class Versions {

  private request: ISPRequest;
  private utils: Utils;
  private baseUrl: string;

  constructor (request: ISPRequest, baseUrl?: string) {
    this.request = request;
    this.utils = new Utils();
    this.baseUrl = baseUrl;
  }

  // GetVersionCollection - for lists items

  /* Documents in libraries */

  public getVersions = (data: IDocumentVersions) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const soapBody: string = this.utils.soapEnvelope(`
      <GetVersions xmlns="http://schemas.microsoft.com/sharepoint/soap/">
        <fileName>${data.fileName}</fileName>
      </GetVersions>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/versions.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return result['soap:Envelope']['soap:Body'][0]
        .GetVersionsResponse[0].GetVersionsResult[0].results[0].result;
    }).then((result) => {
      return result.map((ver) => {
        return ver.$;
      });
    }) as any;
  }

  // Do not work?
  public deleteAllVersions = (data: IDocumentVersions) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const soapBody: string = this.utils.soapEnvelope(`
      <DeleteAllVersions xmlns="http://schemas.microsoft.com/sharepoint/soap/">
        <fileName>${data.fileName}</fileName>
      </DeleteAllVersions>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/versions.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }) as any; // ToDo: results path
  }

  // Do not work?
  public deleteVersion = (data: IDocumentVersion) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const soapBody: string = this.utils.soapEnvelope(`
      <DeleteVersion xmlns="http://schemas.microsoft.com/sharepoint/soap/">
        <fileName>${data.fileName}</fileName>
        <fileVersion>${data.fileVersion}</fileVersion>
      </DeleteVersion>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/versions.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }) as any; // ToDo: results path
  }

  // Do not work?
  public restoreVersion = (data: IDocumentVersion) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const soapBody: string = this.utils.soapEnvelope(`
      <RestoreVersion xmlns="http://schemas.microsoft.com/sharepoint/soap/">
        <fileName>${data.fileName}</fileName>
        <fileVersion>${data.fileVersion}</fileVersion>
      </RestoreVersion>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/versions.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }) as any; // ToDo: results path
  }

  /* Items in lists */

  public getVersionCollection = (data: IItemVersions) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const soapBody: string = this.utils.soapEnvelope(`
      <GetVersionCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/">
        <strlistID>${data.listId}</strlistID>
        <strlistItemID>${data.itemId}</strlistItemID>
        <strFieldName>${data.fieldName}</strFieldName>
      </GetVersionCollection>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/lists.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return result['soap:Envelope']['soap:Body'][0]
        .GetVersionCollectionResponse[0].GetVersionCollectionResult[0];
    }).then((result) => {
      return result.Versions[0].Version.map((v) => v.$);
    }) as any;
  }

}
