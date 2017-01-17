var Handlebars = require('handlebars');

var spf = spf || {};

spf.UPS = function(request) {

    this.modifyUserPropertyByAccountName = function(data) {
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

        soapBody = soapTemplate(data); // .replace(/\\"/g, '"').replace('"<', '<').replace('>"', '>');

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

    return this;

};

module.exports = spf.UPS;