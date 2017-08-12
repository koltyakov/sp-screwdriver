import { ISPRequest } from 'sp-request';
import { Utils } from './../utils';

import { IGetUserProfileByNameProperties, IModifyUserPropertyByAccountNameProperties, IGetUserPropertyByAccountNameProperties,
         IGetUserProfilePropertyForProperties, IGetPropertiesForProperties, ISetSingleValueProfilePropertyProperties,
         ISetMultiValuedProfilePropertyProperties, INewPropData } from './../interfaces/IUps';

export class UPS {

    private request: ISPRequest;
    private utils: Utils;
    private baseUrl: string;

    constructor(request: ISPRequest, baseUrl?: string) {
        this.request = request;
        this.utils = new Utils();
    }

    /* SOAP */

    public getUserProfileByName = (data: IGetUserProfileByNameProperties) => {

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
                    <GetUserProfileByName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">
                        <AccountName>${data.accountName}</AccountName>
                    </GetUserProfileByName>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/UserProfileService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }).then(result => {
            return result['soap:Envelope']['soap:Body'][0]
                .GetUserProfileByNameResponse[0].GetUserProfileByNameResult[0];
        });
    }

    public getUserPropertyByAccountName = (data: IGetUserPropertyByAccountNameProperties) => {

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
                    <GetUserPropertyByAccountName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">
                        <accountName>${data.accountName}</accountName>
                        <propertyName>${data.propertyName}</propertyName>
                    </GetUserPropertyByAccountName>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/UserProfileService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => {
            return this.utils.parseXml(response.body);
        }).then(result => {
            return result['soap:Envelope']['soap:Body'][0]
                .GetUserPropertyByAccountNameResponse[0].GetUserPropertyByAccountNameResult[0];
        });
    }

    public modifyUserPropertyByAccountName = (data: IModifyUserPropertyByAccountNameProperties) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        data.newData = data.newData.map(newData => {
            return <INewPropData>{
                privacy: newData.privacy || 'NotSet',
                isPrivacyChanged: typeof newData.isPrivacyChanged === 'undefined' ? false : newData.isPrivacyChanged,
                isValueChanged: typeof newData.isValueChanged === 'undefined' ? true : newData.isValueChanged
            };
        });

        let soapBody: string = this.utils.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <ModifyUserPropertyByAccountName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">
                        <accountName>${data.accountName}</accountName>
                        <newData>
                            ${
                                data.newData.reduce((res: string, prop) => {
                                    res += `
                                        <PropertyData>
                                            <IsPrivacyChanged>${prop.isPrivacyChanged}</IsPrivacyChanged>
                                            <IsValueChanged>${prop.isValueChanged}</IsValueChanged>
                                            <Name>${prop.name}</Name>
                                            <Privacy>${prop.privacy}</Privacy>
                                            <Values>
                                                ${
                                                    prop.values.reduce((propRes: string, propVal) => {
                                                        propRes += `
                                                            <ValueData>
                                                                <Value xsi:type="xsd:string">${propVal}</Value>
                                                            </ValueData>
                                                        `;
                                                        return propRes;
                                                    }, '')
                                                }
                                            </Values>
                                        </PropertyData>
                                    `;
                                    return res;
                                }, '')
                            }
                        </newData>
                    </ModifyUserPropertyByAccountName>
                </soap:Body>
            </soap:Envelope>
        `);

        let headers: Headers = this.utils.soapHeaders(soapBody);

        return <any>this.request.post(`${data.baseUrl}/_vti_bin/UserProfileService.asmx`, {
            headers,
            body: soapBody,
            json: false
        }).then(response => response.body);
    }

    /* REST */

    public getPropertiesFor = (data: IGetPropertiesForProperties) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let methodUrl = `${data.baseUrl}/_api/sp.userprofiles.peoplemanager` +
            `/getpropertiesfor(` +
                `accountName='${encodeURIComponent(data.accountName)}')`;
        return <any>this.request.get(methodUrl);
    }

    public getUserProfilePropertyFor = (data: IGetUserProfilePropertyForProperties) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let methodUrl = `${data.baseUrl}/_api/sp.userprofiles.peoplemanager` +
            `/getuserprofilepropertyfor(` +
                `accountName='${encodeURIComponent(data.accountName)}',` +
                `propertyname='${data.propertyName}')`;
        return <any>this.request.get(methodUrl);
    }

    /* HTTP */

    public setSingleValueProfileProperty = (data: ISetSingleValueProfilePropertyProperties) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let requestBody: string = this.utils.trimMultiline(`
            <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
                <Actions>
                    <ObjectPath Id="71" ObjectPathId="70" />
                    <Method Name="SetSingleValueProfileProperty" Id="72" ObjectPathId="70">
                        <Parameters>
                            <Parameter Type="String">${data.accountName}</Parameter>
                            <Parameter Type="String">${data.propertyName}</Parameter>
                            <Parameter Type="String">${data.propertyValue}</Parameter>
                        </Parameters>
                    </Method>
                </Actions>
                <ObjectPaths>
                    <Constructor Id="70" TypeId="{cf560d69-0fdb-4489-a216-b6b47adf8ef8}" />
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
                }).then(response => response.body);
            });
    }

    public setMultiValuedProfileProperty = (data: ISetMultiValuedProfilePropertyProperties) => {

        data.baseUrl = data.baseUrl || this.baseUrl;

        if (typeof data.baseUrl === 'undefined') {
            throw new Error('Site URL should be defined');
        }

        let requestBody: string = this.utils.trimMultiline(`
            <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
                <Actions>
                    <ObjectPath Id="82" ObjectPathId="81" />
                    <Method Name="SetMultiValuedProfileProperty" Id="83" ObjectPathId="81">
                        <Parameters>
                            <Parameter Type="String">${data.accountName}</Parameter>
                            <Parameter Type="String">${data.propertyName}</Parameter>
                            <Parameter Type="Array">
                                ${
                                    data.propertyValues.reduce((res: string, propVal) => {
                                        res += `
                                            <Object Type="String">${propVal}</Object>
                                        `;
                                        return res;
                                    }, '')
                                }
                            </Parameter>
                        </Parameters>
                    </Method>
                </Actions>
                <ObjectPaths>
                    <Constructor Id="81" TypeId="{cf560d69-0fdb-4489-a216-b6b47adf8ef8}" />
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
                }).then(response => response.body);
            });
    }

}
