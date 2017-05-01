const Handlebars = require('handlebars');
const util = require('../util');

const Items = function(request) {

    /* HTTP (CSOM) */

    this.setItemProperties = (data) => {
        let sequenceId = 7; // ObjectPathId="6" + 1
        data.properties.forEach(prop => {
            prop.id = sequenceId;
            sequenceId += 1;
        });
        data.updateId = sequenceId;
        data.queryId = sequenceId + 1;

        let requestTemplate = Handlebars.compile(util.trimMultiline(`
            <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
                <Actions>
                    {{#properties}}
                    <Method Name="SetFieldValue" Id="{{ id }}" ObjectPathId="6">
                        <Parameters>
                            <Parameter Type="String">{{ field }}</Parameter>
                            <Parameter Type="String">{{ value }}</Parameter>
                        </Parameters>
                    </Method>
                    {{/properties}}
                    <Method Name="Update" Id="{{ updateId }}" ObjectPathId="6" />
                    <Query Id="{{ queryId }}" ObjectPathId="6">
                        <Query SelectAllProperties="true">
                            <Properties>
                                {{#properties}}
                                <Property Name="{{ field }}" ScalarProperty="true" />
                                {{/properties}}
                            </Properties>
                        </Query>
                    </Query>
                </Actions>
                <ObjectPaths>
                    <StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />
                    <Property Id="2" ParentId="0" Name="Web" />
                    <Method Id="4" ParentId="2" Name="GetList">
                        <Parameters>
                            <Parameter Type="String">{{ listPath }}</Parameter>
                        </Parameters>
                    </Method>
                    <Method Id="6" ParentId="4" Name="GetItemById">
                        <Parameters>
                            <Parameter Type="Number">{{ itemId }}</Parameter>
                        </Parameters>
                    </Method>
                </ObjectPaths>
            </Request>
        `));

        return request.requestDigest(data.baseUrl)
            .then(function(digest) {

                let headers = {};
                let requestBody = '';

                requestBody = requestTemplate(data);

                headers["Accept"] = "*/*";
                headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
                headers["X-Requested-With"] = "XMLHttpRequest";
                headers["Content-Length"] = requestBody.length;
                headers["X-RequestDigest"] = digest;

                return request.post(data.baseUrl + '/_vti_bin/client.svc/ProcessQuery', {
                    headers: headers,
                    body: requestBody,
                    json: false
                });
            });
    };

    return this;

};

module.exports = Items;