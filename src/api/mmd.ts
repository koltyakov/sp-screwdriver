import { ISPRequest } from 'sp-request';
import { Utils } from './../utils';

import {
  IGetTermSetsParams, IGetChildTermsInTermSetParams, IGetChildTermsInTermParams,
  IGetTermsByLabelParams, IGetKeywordTermsByGuidsParams, IAddTermsParams,
  IGetAllTermsParams, ISetTermNameParams, IDeprecateTermsParams,
  ITermSetsResponse, ITerm
} from './../interfaces/IMmd';

export class MMD {

  private request: ISPRequest;
  private utils: Utils;
  private baseUrl: string;

  constructor (request: ISPRequest, baseUrl?: string) {
    this.request = request;
    this.utils = new Utils();
    this.baseUrl = baseUrl;
  }

  /* SOAP */

  public getTermSets = (data: IGetTermSetsParams): Promise<ITermSetsResponse> => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.lcid = data.lcid || 1033;
    data.clientTimeStamp = data.clientTimeStamp || (new Date()).toISOString();
    data.clientVersion = data.clientVersion || 1;

    const soapBody: string = this.utils.soapEnvelope(`
      <GetTermSets xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
        <sharedServiceIds>
          ${
          this.utils.toInnerXmlPackage(`
            <sspIds>
              ${
              [data.sspId].reduce((res: string, sspId) => {
                res += `
                  <sspId>
                    ${sspId}
                  </sspId>
                `;
                return res;
              }, '')
              }
            </sspIds>
          `)
          }
        </sharedServiceIds>
        <termSetIds>
          ${
          this.utils.toInnerXmlPackage(`
            <termSetIds>
                ${
                [data.termSetId].reduce((res: string, termSetId) => {
                  res += `
                    <termSetId>
                      ${termSetId}
                    </termSetId>
                  `;
                  return res;
                }, '')
                }
            </termSetIds>
          `)
          }
        </termSetIds>
        <lcid>${data.lcid}</lcid>
        <clientTimeStamps>
          ${
          this.utils.toInnerXmlPackage(`
            <dateTimes>
              <dateTime>
                ${data.clientTimeStamp}
              </dateTime>
            </dateTimes>
          `)
          }
        </clientTimeStamps>
        <clientVersions>
          ${
          this.utils.toInnerXmlPackage(`
            <versions>
              <version>
                ${data.clientVersion}
              </version>
            </versions>
          `)
          }
        </clientVersions>
      </GetTermSets>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return this.utils.parseXml(
        result['soap:Envelope']['soap:Body'][0]
          .GetTermSetsResponse[0].GetTermSetsResult[0]
      );
    }).then((result) => {
      return result.Container.TermStore;
    }).then(this.mapTermSetFromSoapResponse) as any;
  }

  public getChildTermsInTermSet = (data: IGetChildTermsInTermSetParams): Promise<ITerm[]> => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.lcid = data.lcid || 1033;

    const soapBody: string = this.utils.soapEnvelope(`
      <GetChildTermsInTermSet xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
        <sspId>${data.sspId}</sspId>
        <lcid>${data.lcid}</lcid>
        <termSetId>${data.termSetId}</termSetId>
      </GetChildTermsInTermSet>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return this.utils.parseXml(
        result['soap:Envelope']['soap:Body'][0]
          .GetChildTermsInTermSetResponse[0].GetChildTermsInTermSetResult[0]
      );
    }).then((terms) => {
      return this.mapTermsFromSoapResponse(terms.TermStore.T);
    }) as any;
  }

  public getChildTermsInTerm = (data: IGetChildTermsInTermParams) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.lcid = data.lcid || 1033;

    const soapBody: string = this.utils.soapEnvelope(`
      <GetChildTermsInTerm xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
        <sspId>${data.sspId}</sspId>
        <lcid>${data.lcid}</lcid>
        <termId>${data.termId}</termId>
        <termSetId>${data.termSetId}</termSetId>
      </GetChildTermsInTerm>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return this.utils.parseXml(
        result['soap:Envelope']['soap:Body'][0]
          .GetChildTermsInTermResponse[0].GetChildTermsInTermResult[0]
      );
    }).then((terms) => {
      return this.mapTermsFromSoapResponse(terms.TermStore.T);
    }) as any;
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

    const soapBody: string = this.utils.soapEnvelope(`
      <GetTermsByLabel xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
          <label>${data.label}</label>
          <lcid>${data.lcid}</lcid>
          <matchOption>${data.matchOption}</matchOption>
          <resultCollectionSize>${data.resultCollectionSize}</resultCollectionSize>
          <termIds>
              ${
              this.utils.toInnerXmlPackage(`
                <termIds>
                  ${
                  data.termIds.reduce((res: string, termId) => {
                    res += `
                      <termId>
                        ${termId}
                      </termId>
                    `;
                    return res;
                  }, '')
                  }
                </termIds>
              `)
              }
          </termIds>
          <addIfNotFound>${data.addIfNotFound}</addIfNotFound>
      </GetTermsByLabel>
  `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return this.utils.parseXml(
        result['soap:Envelope']['soap:Body'][0]
          .GetTermsByLabelResponse[0].GetTermsByLabelResult[0]
      );
    }).then((terms) => {
      return this.mapTermsFromSoapResponse(terms.TermStore.T);
    }) as any;
  }

  public getKeywordTermsByGuids = (data: IGetKeywordTermsByGuidsParams) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.lcid = data.lcid || 1033;

    const soapBody: string = this.utils.soapEnvelope(`
      <GetKeywordTermsByGuids xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
        <termIds>
            ${
            this.utils.toInnerXmlPackage(`
              <termIds>
                ${
                data.termIds.reduce((res: string, termId) => {
                  res += `
                    <termId>
                      ${termId}
                    </termId>
                  `;
                  return res;
                }, '')
                }
              </termIds>
            `)
            }
        </termIds>
        <lcid>${data.lcid}</lcid>
      </GetKeywordTermsByGuids>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return this.utils.parseXml(
        result['soap:Envelope']['soap:Body'][0]
          .GetKeywordTermsByGuidsResponse[0].GetKeywordTermsByGuidsResult[0]
      );
    }).then((terms) => {
      return this.mapTermsFromSoapResponse(terms.TermStore.T);
    }) as any;
  }

  public addTerms = (data: IAddTermsParams) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.newTerms = data.newTerms.map((term) => {
      return {
        ...term,
        parentTermId: term.parentTermId || '00000000-0000-0000-0000-000000000000'
      };
    });

    data.lcid = data.lcid || 1033;

    const soapBody: string = this.utils.soapEnvelope(`
      <AddTerms xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">
        <sharedServiceId>${data.sspId}</sharedServiceId>
        <termSetId>${data.termSetId}</termSetId>
        <lcid>${data.lcid}</lcid>
        <newTerms>
          ${
          this.utils.toInnerXmlPackage(`
            <newTerms>
              ${
              data.newTerms.reduce((res: string, newTerm) => {
                res += `<newTerm label="${newTerm.label}" ` +
                  `clientId="1" parentTermId="${newTerm.parentTermId}">` +
                  `</newTerm>`;
                return res;
              }, '')
              }
            </newTerms>
          `)
          }
        </newTerms>
      </AddTerms>
    `);

    const headers: any = this.utils.soapHeaders(soapBody);

    return this.request.post(`${data.baseUrl}/_vti_bin/TaxonomyClientService.asmx`, {
      headers,
      body: soapBody,
      json: false
    }).then((response) => {
      return this.utils.parseXml(response.body);
    }).then((result) => {
      return this.utils.parseXml(
        result['soap:Envelope']['soap:Body'][0]
          .AddTermsResponse[0].AddTermsResult[0]
      );
    }).then((terms) => {
      return this.mapTermsFromSoapResponse(terms.TermStore.T);
    }) as any;
  }

  /* HTTP (CSOM) */

  public getAllTerms = (data: IGetAllTermsParams) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.properties = data.properties || [];

    const requestBody: string = this.utils.trimMultiline(`
      <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
        <Actions>
          <Query Id="78" ObjectPathId="76">
            <Query SelectAllProperties="true">
              <Properties />
            </Query>
            <ChildItemQuery SelectAllProperties="true">
              ${
              (data.properties.length === 0) ?
                `<Properties />` :
                `<Properties>
                  ${
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
                </Properties>`
              }
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
          return result[result.length - 1]._Child_Items_;
        }) as any;
      }) as any;
  }

  public setTermName = (data: ISetTermNameParams) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    const requestBody: string = this.utils.trimMultiline(`
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

  public deprecateTerm = (data: IDeprecateTermsParams) => {

    data.baseUrl = data.baseUrl || this.baseUrl;

    if (typeof data.baseUrl === 'undefined') {
      throw new Error('Site URL should be defined');
    }

    data.deprecate = typeof data.deprecate === 'undefined' ? true : data.deprecate;

    const requestBody: string = this.utils.trimMultiline(`
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

    return this.request.requestDigest(data.baseUrl)
      .then((digest) => {

        if (typeof data.deprecate === 'undefined') {
          data.deprecate = true;
        }

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

  private mapTermSetFromSoapResponse = (soap: any): ITermSetsResponse => {
    return {
      termSet: {
        name: soap[0].TS[0].$.a12,
        id: soap[0].TS[0].$.a9,
        _raw: soap[0].TS[0]
      },
      terms: this.mapTermsFromSoapResponse(soap[0].T)
    };
  }

  private mapTermsFromSoapResponse = (soap: any[] = []): ITerm[] => {
    return soap.map((t) => {
      return {
        name: t.LS[0].TL[0].$.a32,
        id: t.$.a9,
        enableForTagging: t.TMS[0].TM[0].$.a17,
        parentId: t.TMS[0].TM[0].$.a25,
        termSetId: t.TMS[0].TM[0].$.a24,
        _raw: t
      };
    });
  }

}
