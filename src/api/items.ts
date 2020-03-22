import { ISPRequest } from 'sp-request';
import { Utils } from './../utils';

import { ISetItemsPropertiesParams } from './../interfaces/IItems';

export class Items {

  private request: ISPRequest;
  private utils: Utils;
  private baseUrl: string;

  constructor (request: ISPRequest, baseUrl?: string) {
    this.request = request;
    this.utils = new Utils();
    this.baseUrl = baseUrl;
  }

  /* HTTP (CSOM) */

  public setItemProperties = (data: ISetItemsPropertiesParams) => {
    let sequenceId = 7; // ObjectPathId="6" + 1

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.properties.forEach((prop) => {
      prop.id = sequenceId;
      sequenceId += 1;
    });
    data.updateId = sequenceId;
    data.queryId = sequenceId + 1;

    // listPath can be relative
    const relUrl = this.utils.relativeFromAbsoluteUrl(data.baseUrl);
    if (data.listPath.indexOf(relUrl) === -1) {
      data.listPath = `${relUrl}/${data.listPath}`.replace(/\/\//g, '/');
    }

    const requestBody: string = this.utils.trimMultiline(`
      <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
        <Actions>
          ${
          data.properties.reduce((res: string, property) => {
            res += `
              <Method Name="SetFieldValue" Id="${property.id}" ObjectPathId="6">
                <Parameters>
                  <Parameter Type="String">${property.field}</Parameter>
                  <Parameter Type="String">${property.value}</Parameter>
                </Parameters>
              </Method>
            `;
            return res;
          }, '')
          }
          <Method Name="Update" Id="${data.updateId}" ObjectPathId="6" />
          <Query Id="${data.queryId}" ObjectPathId="6">
            <Query SelectAllProperties="true">
              <Properties>
                ${
                data.properties.reduce((res: string, property) => {
                  res += `
                    <Property Name="${property.field}" ScalarProperty="true" />
                  `;
                  return res;
                }, '')
                }
              </Properties>
            </Query>
          </Query>
        </Actions>
        <ObjectPaths>
          <StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />
          <Property Id="2" ParentId="0" Name="Web" />
          <Method Id="4" ParentId="2" Name="GetList">
            <Parameters>
              <Parameter Type="String">${data.listPath}</Parameter>
            </Parameters>
          </Method>
          <Method Id="6" ParentId="4" Name="GetItemById">
            <Parameters>
              <Parameter Type="Number">${data.itemId}</Parameter>
            </Parameters>
          </Method>
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
          return result[result.length - 1];
        });

      }) as any;
  }

}
