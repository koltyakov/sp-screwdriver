var Handlebars = require('handlebars');

var spf = spf || {};

spf.UPS = function(request) {

    /* SOAP */

    this.getUserProfileByName = (data) => {
        var headers = {};
        var soapBody = '';
        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<GetUserProfileByName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">' +
                        '<AccountName>{{ accountName }}</AccountName>' +
                    '</GetUserProfileByName>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/UserProfileService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.getUserPropertyByAccountName = (data) => {
        var headers = {};
        var soapBody = '';
        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<GetUserPropertyByAccountName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">' +
                        '<accountName>{{ accountName }}</accountName>' +
                        '<propertyName>{{ propertyName }}</propertyName>' +
                    '</GetUserPropertyByAccountName>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/UserProfileService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.modifyUserPropertyByAccountName = (data) => {
        var headers = {};
        var soapBody = '';
        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<ModifyUserPropertyByAccountName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">' +
                        '<accountName>{{ accountName }}</accountName>' +
                        '<newData>' +
                            '{{#newData}}' +
                            '<PropertyData>' +
                                '<IsPrivacyChanged>{{ isPrivacyChanged }}</IsPrivacyChanged>' +
                                '<IsValueChanged>{{ isValueChanged }}</IsValueChanged>' +
                                '<Name>{{ name }}</Name>' +
                                '<Privacy>{{ privacy }}</Privacy>' +
                                '<Values>' +
                                    '{{#values}}' +
                                    '<ValueData>' +
                                        '<Value xsi:type="xsd:string">{{ this }}</Value>' +
                                    '</ValueData>' +
                                    '{{/values}}' +
                                '</Values>' +
                            '</PropertyData>' +
                            '{{/newData}}' +
                        '</newData>' +
                    '</ModifyUserPropertyByAccountName>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        data.newData = (data.newData || []).map(function(data) {
            if (typeof data.value !== "undefined" && typeof data.values === "undefined") {
                data.values = [data.value];
            }
            if (typeof data.privacy) {
                data.privacy = 'NotSet';
            }
            if (typeof data.isPrivacyChanged) {
                data.isPrivacyChanged = false;
            }
            if (typeof data.values !== "undefined" && typeof data.isValueChanged === "undefined") {
                data.isValueChanged = true;
            }
            return data;
        });

        soapBody = soapTemplate(data);

        headers["SOAPAction"] = "http://microsoft.com/webservices/SharePointPortalServer/UserProfileService/ModifyUserPropertyByAccountName";
        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/UserProfileService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    /* REST */

    this.getPropertiesFor = (data) => {
        var methodUrl = `${data.baseUrl}/_api/sp.userprofiles.peoplemanager` +
            `/getpropertiesfor(` +
                `accountName='${encodeURIComponent(data.accountName)}')`;
        return request.get(methodUrl);
    };

    this.getUserProfilePropertyFor = (data) => {
        var methodUrl = `${data.baseUrl}/_api/sp.userprofiles.peoplemanager` +
            `/getuserprofilepropertyfor(` +
                `accountName='${encodeURIComponent(data.accountName)}',` +
                `propertyname='${data.propertyName}')`;
        return request.get(methodUrl);
    };

    /* HTTP */

    this.setSingleValueProfileProperty = (data) => {
        var requestTemplate = Handlebars.compile(`
            <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
                <Actions>
                    <ObjectPath Id="71" ObjectPathId="70" />
                    <Method Name="SetSingleValueProfileProperty" Id="72" ObjectPathId="70">
                        <Parameters>
                            <Parameter Type="String">{{ accountName }}</Parameter>
                            <Parameter Type="String">{{ propertyName }}</Parameter>
                            <Parameter Type="String">{{ propertyValue }}</Parameter>
                        </Parameters>
                    </Method>
                </Actions>
                <ObjectPaths>
                    <Constructor Id="70" TypeId="{cf560d69-0fdb-4489-a216-b6b47adf8ef8}" />
                </ObjectPaths>
            </Request>
        `);

        return request.requestDigest(data.baseUrl)
            .then(function(digest) {

                var headers = {};
                var requestBody = '';

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

    this.setMultiValuedProfileProperty = (data) => {
        var requestTemplate = Handlebars.compile(`
            <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">
                <Actions>
                    <ObjectPath Id="82" ObjectPathId="81" />
                    <Method Name="SetMultiValuedProfileProperty" Id="83" ObjectPathId="81">
                        <Parameters>
                            <Parameter Type="String">{{ accountName }}</Parameter>
                            <Parameter Type="String">{{ propertyName }}</Parameter>
                            <Parameter Type="Array">
                                {{#propertyValues}}
                                <Object Type="String">{{ this }}</Object>
                                {{/propertyValues}}
                            </Parameter>
                        </Parameters>
                    </Method>
                </Actions>
                <ObjectPaths>
                    <Constructor Id="81" TypeId="{cf560d69-0fdb-4489-a216-b6b47adf8ef8}" />
                </ObjectPaths>
            </Request>
        `);

        return request.requestDigest(data.baseUrl)
            .then(function(digest) {

                var headers = {};
                var requestBody = '';

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

module.exports = spf.UPS;