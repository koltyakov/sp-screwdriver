var Handlebars = require('handlebars');

var spf = spf || {};

spf.MMD = function(request) {

    /* SOAP */

    this.getTermSets = function(data) {
        var headers = {};
        var soapBody = '';

        data.lcid = data.lcid || 1033;
        data.version = data.version || 1;

        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<GetTermSets xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">' +
                        '<sharedServiceIds>' +
                            '&lt;sspIds&gt;' +
                                '{{#sspIds}}' +
                                    '&lt;sspId&gt;' +
                                        '{{ this }}' +
                                    '&lt;/sspId&gt;' +
                                '{{/sspIds}}' +
                            '&lt;/sspIds&gt;' +
                        '</sharedServiceIds>' +
                        '<termSetIds>' +
                            '&lt;termSetIds&gt;' +
                                '{{#termSetIds}}' +
                                    '&lt;termSetId&gt;' +
                                        '{{ this }}' +
                                    '&lt;/termSetId&gt;' +
                                '{{#termSetIds}}' +
                            '&lt;/termSetIds&gt;' +
                        '</termSetIds>' +
                        '<lcid>{{ lcid }}</lcid>' +
                        '<clientTimeStamps>' +
                            '&lt;dateTimes&gt;&lt;dateTime&gt;' +
                                '{{ clientTimeStamp }}' +
                            '&lt;/dateTime&gt;&lt;/dateTimes&gt;' +
                        '</clientTimeStamps>' +
                        '<clientVersions>' +
                            '&lt;versions&gt;&lt;version&gt;' +
                                '{{ clientVersion }}' +
                            '&lt;/version&gt;&lt;/versions&gt;' +
                        '</clientVersions>' +
                    '</GetTermSets>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        if (typeof data.sspId !== "undefined" && typeof data.sspIds === "undefined") {
            data.sspIds = [data.sspId];
        }
        if (typeof data.termSetId !== "undefined" && typeof data.termSetIds === "undefined") {
            data.termSetIds = [data.termSetId];
        }


        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/TaxonomyClientService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.getChildTermsInTermSet = function(data) {
        var headers = {};
        var soapBody = '';

        data.lcid = data.lcid || 1033;
        data.version = data.version || 1;

        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<GetChildTermsInTermSet xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">' +
                        '<sspId>{{ sspId }}</sspId>' +
                        '<lcid>{{ lcid }}</lcid>' +
                        '<termSetId>{{ termSetId }}</termSetId>' +
                    '</GetChildTermsInTermSet>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/TaxonomyClientService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.getChildTermsInTerm = function(data) {
        var headers = {};
        var soapBody = '';

        data.lcid = data.lcid || 1033;
        data.version = data.version || 1;

        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<GetChildTermsInTerm xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">' +
                        '<sspId>{{ sspId }}</sspId>' +
                        '<lcid>{{ lcid }}</lcid>' +
                        '<termId>{{ termId }}</termId>' +
                        '<termSetId>{{ termSetId }}</termSetId>' +
                    '</GetChildTermsInTerm>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/TaxonomyClientService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.getTermsByLabel = function(data) {
        var headers = {};
        var soapBody = '';

        data.lcid = data.lcid || 1033;
        data.version = data.version || 1;

        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<GetTermsByLabel xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">' +
                        '<label>{{ label }}</label>' +
                        '<lcid>{{ lcid }}</lcid>' +
                        '<matchOption>{{ matchOption }}</matchOption>' +
                        '<resultCollectionSize>{{ resultCollectionSize }}</resultCollectionSize>' +
                        '<termIds>' +
                            '&lt;termIds&gt;' +
                                '{{#termIds}}' +
                                    '&lt;termId&gt;' +
                                        '{{ this }}' +
                                    '&lt;/termId&gt;' +
                                '{{/termIds}}' +
                            '&lt;/termIds&gt;' +
                        '</termIds>' +
                        '<addIfNotFound>{{ addIfNotFound }}</addIfNotFound>' +
                    '</GetTermsByLabel>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        data.matchOption = data.matchOption || "ExactMatch"; // or StartsWith
        data.resultCollectionSize = data.resultCollectionSize || 25;
        if (typeof data.addIfNotFound === "undefined") {
            data.addIfNotFound = false;
        }
        if (typeof data.termId !== "undefined" && typeof data.termIds === "undefined") {
            data.termIds = [data.termId];
        }
        data.termIds = data.termIds || [];

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/TaxonomyClientService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.getKeywordTermsByGuids = function(data) {
        var headers = {};
        var soapBody = '';

        data.lcid = data.lcid || 1033;
        data.version = data.version || 1;

        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<GetKeywordTermsByGuids xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">' +
                        '<termIds>' +
                            '&lt;termIds&gt;' +
                                '{{#termIds}}' +
                                    '&lt;termId&gt;' +
                                        '{{ this }}' +
                                    '&lt;/termId&gt;' +
                                '{{/termIds}}' +
                            '&lt;/termIds&gt;' +
                        '</termIds>' +
                        '<lcid>{{ lcid }}</lcid>' +
                    '</GetKeywordTermsByGuids>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        if (typeof data.termId !== "undefined" && typeof data.termIds === "undefined") {
            data.termIds = [data.termId];
        }

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/TaxonomyClientService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    this.addTerms = function(data) {
        var headers = {};
        var soapBody = '';

        data.lcid = data.lcid || 1033;
        data.version = data.version || 1;

        var soapTemplate = Handlebars.compile(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                    '<AddTerms xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">' +
                        '<sharedServiceId>{{ sspId }}</sharedServiceId>' +
                        '<termSetId>{{ termSetId }}</termSetId>' +
                        '<lcid>{{ lcid }}</lcid>' +
                        '<newTerms>' +
                            '{{ newTerms }}' +
                        '</newTerms>' +
                    '</AddTerms>' +
                '</soap:Body>' +
            '</soap:Envelope>'
        );

        soapBody = soapTemplate(data);

        headers["Accept"] = "application/xml, text/xml, */*; q=0.01";
        headers["Content-Type"] = "text/xml;charset=\"UTF-8\"";
        headers["X-Requested-With"] = "XMLHttpRequest";
        headers["Content-Length"] = soapBody.length;

        return request.post(data.baseUrl + '/_vti_bin/TaxonomyClientService.asmx', {
            headers: headers,
            body: soapBody,
            json: false
        });
    };

    /* HTTP */

    this.getAllTerms = function(data) {

        var requestTemplate = Handlebars.compile(
            '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName="Javascript Library">' +
                '<Actions>' +
                    '<Query Id="78" ObjectPathId="76">' +
                        '<Query SelectAllProperties="true">' +
                            '<Properties />' +
                        '</Query>' +
                        '<ChildItemQuery SelectAllProperties="true">' +
                            '<Properties />' +
                        '</ChildItemQuery>' +
                    '</Query>' +
                '</Actions>' +
                '<ObjectPaths>' +
                    '<StaticMethod Id="65" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" />' +
                    '<Property Id="68" ParentId="65" Name="TermStores" />' +
                    '<Method Id="70" ParentId="68" Name="GetByName">' +
                        '<Parameters>' +
                            '<Parameter Type="String">{{ serviceName }}</Parameter>' +
                        '</Parameters>' +
                    '</Method>' +
                    '<Method Id="73" ParentId="70" Name="GetTermSet">' +
                        '<Parameters>' +
                            '<Parameter Type="String">{{ termSetId }}</Parameter>' +
                        '</Parameters>' +
                    '</Method>' +
                    '<Method Id="76" ParentId="73" Name="GetAllTerms" />' +
                '</ObjectPaths>' +
            '</Request>'
        );

        return request.requestDigest(data.baseUrl)
            .then(function(digest) {

                var headers = {};
                var srequestBody = '';

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

module.exports = spf.MMD;