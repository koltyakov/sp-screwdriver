import { ISPRequest } from 'sp-request';
import { Utils } from './../utils';

import { IDocumentVersions, IDocumentVersion, IItemVersions } from './../interfaces/IVersions';

export class Versions {

    private request: ISPRequest;
    private utils: Utils;
    private baseUrl: string;

    constructor(request: ISPRequest, baseUrl?: string) {
        this.request = request;
        this.utils = new Utils();
    }

    // GetVersionCollection - for lists items

    /* Documents in libraries */

    public getVersions = (data: IDocumentVersions) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetVersions xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <fileName>${data.fileName}</fileName>
                    </GetVersions>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/versions.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }).then(result => {
            return result['soap:Envelope']['soap:Body'][0]
                .GetVersionsResponse[0].GetVersionsResult[0].results[0].result;
        });
    }

    public deleteAllVersions = (data: IDocumentVersions) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <DeleteAllVersions xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <fileName>${data.fileName}</fileName>
                    </DeleteAllVersions>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/versions.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }); // ToDo: results path
    }

    public deleteVersion = (data: IDocumentVersion) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <DeleteVersion xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <fileName>${data.fileName}</fileName>
                        <fileVersion>${data.fileVersion}</fileVersion>
                    </DeleteVersion>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/versions.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }); // ToDo: results path
    }

    public restoreVersion = (data: IDocumentVersion) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <RestoreVersion xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <fileName>${data.fileName}</fileName>
                        <fileVersion>${data.fileVersion}</fileVersion>
                    </RestoreVersion>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/versions.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }); // ToDo: results path
    }

    /* Items in lists */

    public getVersionCollection = (data: IItemVersions) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetVersionCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <strlistID>${data.listId}</strlistID>
                        <strlistItemID>${data.itemId}</strlistItemID>
                        <strFieldName>${data.fieldName}</strFieldName>
                    </GetVersionCollection>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/lists.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }); // ToDo: results path
    }

}
