const Handlebars = require('handlebars');
const util = require('../util');

const Versions = function(request) {

    // GetVersionCollection - for lists items

    /* Documents in libraries */

    this.getVersions = (data) => {
        let headers = {};
        let soapBody = '';
        let soapTemplate = Handlebars.compile(util.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetVersions xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <fileName>{{ fileName }}</fileName>
                    </GetVersions>
                </soap:Body>
            </soap:Envelope>
        `));

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/versions.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.deleteAllVersions = (data) => {
        let headers = {};
        let soapBody = '';
        let soapTemplate = Handlebars.compile(util.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <DeleteAllVersions xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <fileName>{{ fileName }}</fileName>
                    </DeleteAllVersions>
                </soap:Body>
            </soap:Envelope>
        `));

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/versions.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.deleteVersion = (data) => {
        let headers = {};
        let soapBody = '';
        let soapTemplate = Handlebars.compile(util.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <DeleteVersion xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <fileName>{{ fileName }}</fileName>
                        <fileVersion>{{ fileVersion }}</fileVersion>
                    </DeleteVersion>
                </soap:Body>
            </soap:Envelope>
        `));

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/versions.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.restoreVersion = (data) => {
        let headers = {};
        let soapBody = '';
        let soapTemplate = Handlebars.compile(util.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <RestoreVersion xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <fileName>{{ fileName }}</fileName>
                        <fileVersion>{{ fileVersion }}</fileVersion>
                    </RestoreVersion>
                </soap:Body>
            </soap:Envelope>
        `));

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/versions.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    /* Items in lists */

    this.getVersionCollection = (data) => {
        let headers = {};
        let soapBody = '';
        let soapTemplate = Handlebars.compile(util.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <GetVersionCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <strlistID>{{ listId }}</strlistID>
                        <strlistItemID>{{ itemId }}</strlistItemID>
                        <strFieldName>{{ fieldName }}</strFieldName>
                    </GetVersionCollection>
                </soap:Body>
            </soap:Envelope>
        `));

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/lists.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    return this;

};

module.exports = Versions;