import { ISPRequest } from 'sp-request';
import { Utils } from './../utils';

import {
  IGetUserProfileByNameProperties, IModifyUserPropertyByAccountNameProperties, IGetUserPropertyByAccountNameProperties,
  IGetUserProfilePropertyForProperties, IGetPropertiesForProperties, ISetSingleValueProfilePropertyProperties,
  ISetMultiValuedProfilePropertyProperties, INewPropData,
  IUserProp
} from './../interfaces/IUps';

export class UPS {

  private request: ISPRequest;
  private utils: Utils;
  private baseUrl: string;

  constructor (request: ISPRequest, baseUrl?: string) {
    this.request = request;
    this.utils = new Utils();
    this.baseUrl = baseUrl;
  }

  /* SOAP */

  public getUserProfileByName = (data: IGetUserProfileByNameProperties): Promise<IUserProp> => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const soapBody: string = this.utils.soapEnvelope(`
      <GetUserProfileByName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">
        <AccountName>${data.accountName}</AccountName>
      </GetUserProfileByName>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/UserProfileService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return result['soap:Envelope']['soap:Body'][0]
        .GetUserProfileByNameResponse[0].GetUserProfileByNameResult[0];
    }).then((props) => {
      return props.PropertyData.map(this.mapUserPropertiesFromSoapResponse);
    }) as any;
  }

  public getUserPropertyByAccountName = (data: IGetUserPropertyByAccountNameProperties) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const soapBody: string = this.utils.soapEnvelope(`
      <GetUserPropertyByAccountName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">
        <accountName>${data.accountName}</accountName>
        <propertyName>${data.propertyName}</propertyName>
      </GetUserPropertyByAccountName>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/UserProfileService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return result['soap:Envelope']['soap:Body'][0]
        .GetUserPropertyByAccountNameResponse[0].GetUserPropertyByAccountNameResult[0];
    }).then((props) => {
      return this.mapUserPropertiesFromSoapResponse(props);
    }) as any;
  }

  public modifyUserPropertyByAccountName = (data: IModifyUserPropertyByAccountNameProperties) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.newData = data.newData.map((newData) => {
      const r: INewPropData = {
        ...newData,
        privacy: newData.privacy || 'NotSet',
        isPrivacyChanged: typeof newData.isPrivacyChanged === 'undefined' ? false : newData.isPrivacyChanged,
        isValueChanged: typeof newData.isValueChanged === 'undefined' ? true : newData.isValueChanged
      };
      return r;
    });

    const soapBody: string = this.utils.soapEnvelope(`
      <ModifyUserPropertyByAccountName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">
          <accountName>${data.accountName}</accountName>
          <newData>
              ${
              this.utils.toInnerXmlPackage(data.newData.reduce((res: string, prop) => {
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
              }, ''))
              }
          </newData>
      </ModifyUserPropertyByAccountName>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/UserProfileService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return result['soap:Envelope']['soap:Body'][0]
        .ModifyUserPropertyByAccountNameResponse;
    }) as any;
  }

  /* REST */

  public getPropertiesFor = (data: IGetPropertiesForProperties) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const methodUrl = `${data.baseUrl}/_api/sp.userprofiles.peoplemanager` +
      `/getpropertiesfor(` +
      `accountName='${encodeURIComponent(data.accountName)}')`;
    return this.request.get(methodUrl)
      .then((response) => response.body) as any;
  }

  public getUserProfilePropertyFor = (data: IGetUserProfilePropertyForProperties) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const methodUrl = `${data.baseUrl}/_api/sp.userprofiles.peoplemanager` +
      `/getuserprofilepropertyfor(` +
      `accountName='${encodeURIComponent(data.accountName)}',` +
      `propertyname='${data.propertyName}')`;
    return this.request.get(methodUrl)
      .then((response) => response.body.d) as any;
  }

  /* HTTP */

  public setSingleValueProfileProperty = (data: ISetSingleValueProfilePropertyProperties) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const requestBody: string = this.utils.trimMultiline(`
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

    return this.request.requestDigest(data.baseUrl)
      .then((digest) => {

        const headers: any = this.utils.csomHeaders(requestBody, digest);

        return this.request.post(`${data.baseUrl}/_vti_bin/client.svc/ProcessQuery`, {
          headers,
          body: requestBody,
          json: false
        }).then((response) => {
          const result: any = JSON.parse(response.body);
          if (result[0].ErrorInfo !== null) {
            throw new Error(JSON.stringify(result[0].ErrorInfo));
          }
          return true;
        }) as any;
      }) as any;
  }

  public setMultiValuedProfileProperty = (data: ISetMultiValuedProfilePropertyProperties) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const requestBody: string = this.utils.trimMultiline(`
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

    return this.request.requestDigest(data.baseUrl)
      .then((digest) => {

        const headers: any = this.utils.csomHeaders(requestBody, digest);

        return this.request.post(`${data.baseUrl}/_vti_bin/client.svc/ProcessQuery`, {
          headers,
          body: requestBody,
          json: false
        }).then((response) => {
          const result: any = JSON.parse(response.body);
          if (result[0].ErrorInfo !== null) {
            throw new Error(JSON.stringify(result[0].ErrorInfo));
          }
          return true;
        }) as any;
      }) as any;
  }

  // Data mapping

  private mapUserPropertiesFromSoapResponse = (prop: any): IUserProp => {
    return {
      name: prop.Name[0],
      values: prop.Values[0] !== '' ? prop.Values[0].ValueData.map((v) => v._).filter((v) => v !== null) : null,
      privacy: prop.Privacy[0],
      isPrivacyChanged: prop.IsPrivacyChanged[0] === 'true' ? true : false,
      isValueChanged: prop.IsValueChanged[0] === 'true' ? true : false
    };
  }

}
