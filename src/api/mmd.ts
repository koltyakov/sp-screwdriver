import { ISPRequest } from 'sp-request';
import { Utils } from './../utils';

import { IGetTermSetsParams, IGetChildTermsInTermSetParams, IGetChildTermsInTermParams,
         IGetTermsByLabelParams, IGetKeywordTermsByGuidsParams, IAddTermsParams,
         IGetAllTermsParams, ISetTermNameParams, IDeprecateTermsParams } from './../interfaces/IMmd';

export class MMD {

    private request: ISPRequest;
    private utils: Utils;
    private baseUrl: string;

    constructor(request: ISPRequest, baseUrl?: string) {
        this.request = request;
        this.utils = new Utils();
    }

    /* SOAP */

    public getTermSets = (data: IGetTermSetsParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.lcid = data.lcid || 1033;
        data.clientTimeStamp = data.clientTimeStamp || (new Date()).toISOString();
        data.clientVersion = data.clientVersion || 1;

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetTermSets xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
                        <sharedServiceIds>
                            &lt;sspIds&gt;
                                ${
                                    data.sspIds.reduce((res: string, sspId) => {
                                        res += `
                                            &lt;sspId&gt;
                                                ${sspId}
                                            &lt;/sspId&gt;
                                        `;
                                        return res;
                                    }, '')
                                }
                            &lt;/sspIds&gt;
                        </sharedServiceIds>
                        <termSetIds>
                            &lt;termSetIds&gt;
                                ${
                                    data.termSetIds.reduce((res: string, termSetId) => {
                                        res += `
                                            &lt;termSetId&gt;
                                                ${termSetId}
                                            &lt;/termSetId&gt;
                                        `;
                                        return res;
                                    }, '')
                                }
                            &lt;/termSetIds&gt;
                        </termSetIds>
                        <lcid>${data.lcid}</lcid>
                        <clientTimeStamps>
                            &lt;dateTimes&gt;&lt;dateTime&gt;
                                ${data.clientTimeStamp}
                            &lt;/dateTime&gt;&lt;/dateTimes&gt;
                        </clientTimeStamps>
                        <clientVersions>
                            &lt;versions&gt;&lt;version&gt;
                                ${data.clientVersion}
                            &lt;/version&gt;&lt;/versions&gt;
                        </clientVersions>
                    </GetTermSets>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }).then(result => {
            return result['soap:Envelope']['soap:Body'][0].GetTermSetsResponse[0];
        });
    }

    public getChildTermsInTermSet = (data: IGetChildTermsInTermSetParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.lcid = data.lcid || 1033;

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetChildTermsInTermSet xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
                        <sspId>${data.sspId}</sspId>
                        <lcid>${data.lcid}</lcid>
                        <termSetId>${data.termSetId}</termSetId>
                    </GetChildTermsInTermSet>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }).then(result => {
            return this.utils.parseXml(
                result['soap:Envelope']['soap:Body'][0]
                    .GetChildTermsInTermSetResponse[0].GetChildTermsInTermSetResult[0]
            );
        }).then(terms => terms.TermStore);
    }

    public getChildTermsInTerm = (data: IGetChildTermsInTermParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.lcid = data.lcid || 1033;

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetChildTermsInTerm xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
                        <sspId>${data.sspId}</sspId>
                        <lcid>${data.lcid}</lcid>
                        <termId>${data.termId}</termId>
                        <termSetId>${data.termSetId}</termSetId>
                    </GetChildTermsInTerm>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }).then(result => {
            return this.utils.parseXml(
                result['soap:Envelope']['soap:Body'][0]
                    .GetChildTermsInTermResponse[0].GetChildTermsInTermResult[0]
            );
        }).then(terms => terms.TermStore);
    }

    public getTermsByLabel = (data: IGetTermsByLabelParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.lcid = data.lcid || 1033;

        data.matchOption = data.matchOption || 'ExactMatch';
        data.resultCollectionSize = data.resultCollectionSize || 25;
        data.addIfNotFound = typeof data.addIfNotFound === 'undefined' ? false : data.addIfNotFound;

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetTermsByLabel xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
                        <label>${data.label}</label>
                        <lcid>${data.lcid}</lcid>
                        <matchOption>${data.matchOption}</matchOption>
                        <resultCollectionSize>${data.resultCollectionSize}</resultCollectionSize>
                        <termIds>
                            &lt;termIds&gt;
                                ${
                                    data.termIds.reduce((res: string, termId) => {
                                        res += `
                                            &lt;termId&gt;
                                                ${termId}
                                            &lt;/termId&gt;
                                        `;
                                        return res;
                                    }, '')
                                }
                            &lt;/termIds&gt;
                        </termIds>
                        <addIfNotFound>${data.addIfNotFound}</addIfNotFound>
                    </GetTermsByLabel>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }).then(result => {
            return this.utils.parseXml(
                result['soap:Envelope']['soap:Body'][0]
                    .GetTermsByLabelResponse[0].GetTermsByLabelResult[0]
            );
        }).then(terms => terms.TermStore);
    }

    public getKeywordTermsByGuids = (data: IGetKeywordTermsByGuidsParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.lcid = data.lcid || 1033;

        let soapBody = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetKeywordTermsByGuids xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
                        <termIds>
                            &lt;termIds&gt;
                                ${
                                    data.termIds.reduce((res: string, termId) => {
                                        res += `
                                            &lt;termId&gt;
                                                ${termId}
                                            &lt;/termId&gt;
                                        `;
                                        return res;
                                    }, '')
                                }
                            &lt;/termIds&gt;
                        </termIds>
                        <lcid>${data.lcid}</lcid>
                    </GetKeywordTermsByGuids>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }).then(result => {
            return this.utils.parseXml(
                result['soap:Envelope']['soap:Body'][0]
                    .GetKeywordTermsByGuidsResponse[0].GetKeywordTermsByGuidsResult[0]
            );
        }).then(terms => terms.TermStore);
    }

    public addTerms = (data: IAddTermsParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.lcid = data.lcid || 1033;

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <AddTerms xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
                        <sharedServiceId>${data.sspId}</sharedServiceId>
                        <termSetId>${data.termSetId}</termSetId>
                        <lcid>${data.lcid}</lcid>
                        <newTerms>
                            <newTerms>
                                ${
                                    data.newTerms.reduce((res: string, newTerm) => {
                                        res += `
                                            <newTerm label="${newTerm.label}" clientId="1" parentTermId="${newTerm.parentTermId}"></newTerm>
                                        `;
                                        return res;
                                    }, '')
                                }
                            </newTerms>
                        </newTerms>
                    </AddTerms>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        });
    }

    /* HTTP (CSOM) */

    public getAllTerms = (data: IGetAllTermsParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.properties = data.properties || [];

        let requestBody: string = this.utils.trimMultiline(`
            <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
                <Actions>
                    <Query Id="78" ObjectPathId="76">
                        <Query SelectAllProperties="true">
                            <Properties />
                        </Query>
                        <ChildItemQuery SelectAllProperties="true">
                            <Properties>
                                ${
                                    (data.properties.length === 0) ?
                                        `<Properties />` :
                                        data.properties.reduce((res: string, propName) => {
                                            if (propName === 'Parent') {
                                                res += `
                                                    <Property Name="Parent">
                                                        <Query SelectAllProperties="false">
                                                            <Properties>
                                                                <Property Name="Id" SelectAll="true" />
                                                            </Properties>
                                                        </Query>
                                                    </Property>
                                                `;
                                            } else {
                                                res += `<Property Name="${propName}" SelectAll="true" />`;
                                            }
                                            return res;
                                        }, '')
                                }
                            </Properties>
                        </ChildItemQuery>
                    </Query>
                </Actions>
                <ObjectPaths>
                    <StaticMethod Id="65" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" />
                    <Property Id="68" ParentId="65" Name="TermStores" />
                    <Method Id="70" ParentId="68" Name="GetByName">
                        <Parameters>
                            <Parameter Type="String">${data.serviceName}</Parameter>
                        </Parameters>
                    </Method>
                    <Method Id="73" ParentId="70" Name="GetTermSet">
                        <Parameters>
                            <Parameter Type="String">${data.termSetId}</Parameter>
                        </Parameters>
                    </Method>
                    <Method Id="76" ParentId="73" Name="GetAllTerms" />
                </ObjectPaths>
            </Request>
        `);

        return <any>this.request.requestDigest(data.baseUrl)
            .then(digest => {

                let headers: Headers = this.utils.csomHeaders(requestBody, digest);

                return <any>this.request.post(`${data.baseUrl}/_vti_bin/client.svc/ProcessQuery`, {
                    headers,
                    body: requestBody,
                    json: false
                }).then(response => {
                    let results = JSON.parse(response.body);
                    return results[results.length - 1]._Child_Items_;
                });
            });
    }

    public setTermName = (data: ISetTermNameParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let requestBody: string = this.utils.trimMultiline(`
            <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
                <Actions>
                    <SetProperty Id="166" ObjectPathId="157" Name="Name">
                        <Parameter Type="String">${data.newName}</Parameter>
                    </SetProperty>
                </Actions>
                <ObjectPaths>
                    <StaticMethod Id="146" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" />
                    <Property Id="149" ParentId="146" Name="TermStores" />
                    <Method Id="151" ParentId="149" Name="GetByName">
                        <Parameters>
                            <Parameter Type="String">${data.serviceName}</Parameter>
                        </Parameters>
                    </Method>
                    <Method Id="154" ParentId="151" Name="GetTermSet">
                        <Parameters>
                            <Parameter Type="String">${data.termSetId}</Parameter>
                        </Parameters>
                    </Method>
                    <Method Id="157" ParentId="154" Name="GetTerm">
                        <Parameters>
                            <Parameter Type="String">${data.termId}</Parameter>
                        </Parameters>
                    </Method>
                </ObjectPaths>
            </Request>
        `);

        return <any>this.request.requestDigest(data.baseUrl)
            .then(digest => {

                let headers: Headers = this.utils.csomHeaders(requestBody, digest);

                return <any>this.request.post(`${data.baseUrl}/_vti_bin/client.svc/ProcessQuery`, {
                    headers,
                    body: requestBody,
                    json: false
                }).then(response => {
                    return response.body;
                });
            });
    }

    public deprecateTerm = (data: IDeprecateTermsParams) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.deprecate = typeof data.deprecate === 'undefined' ? true : data.deprecate;

        let requestBody = this.utils.trimMultiline(`
            <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
                <Actions>
                    <Method Name="Deprecate" Id="41" ObjectPathId="32">
                        <Parameters>
                            <Parameter Type="Boolean">${data.deprecate}</Parameter>
                        </Parameters>
                    </Method>
                </Actions>
                <ObjectPaths>
                    <StaticMethod Id="21" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" />
                    <Property Id="24" ParentId="21" Name="TermStores" />
                    <Method Id="26" ParentId="24" Name="GetByName">
                        <Parameters>
                            <Parameter Type="String">${data.serviceName}</Parameter>
                        </Parameters>
                    </Method>
                    <Method Id="29" ParentId="26" Name="GetTermSet">
                        <Parameters>
                            <Parameter Type="String">${data.termSetId}</Parameter>
                        </Parameters>
                    </Method>
                    <Method Id="32" ParentId="29" Name="GetTerm">
                        <Parameters>
                            <Parameter Type="String">${data.termId}</Parameter>
                        </Parameters>
                    </Method>
                </ObjectPaths>
            </Request>
        `);

        return <any>this.request.requestDigest(data.baseUrl)
            .then(function(digest) {

                if (typeof data.deprecate === 'undefined') {
                    data.deprecate = true;
                }

                let headers: Headers = this.utils.csomHeaders(requestBody, digest);

                return <any>this.request.post(`${data.baseUrl}/_vti_bin/client.svc/ProcessQuery`, {
                    headers,
                    body: requestBody,
                    json: false
                }).then(response => {
                    return response.body;
                });
            });
    }

}
