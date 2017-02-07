var path = require('path');
var cpass = new (require('cpass'));
var parseString = require('xml2js').parseString;

var Screwdriver = require(__dirname + "/../src/index");

var configPath = path.join(__dirname + "/config/_private.conf.json");
var config = require(configPath);
var context = config.context;
if (context.password) {
    context.password = cpass.decode(context.password);
}

var screw = new Screwdriver(context);

// ================================

var getTermSets = function() {

    var data = {
        baseUrl: context.siteUrl,
        sspId: config.mmd.sspId,
        // sspIds: [config.mmd.sspId],
        termSetId: config.mmd.termSetId,
        // termSetIds: [config.mmd.termSetId],
        lcid: config.mmd.lcid,
        clientTimeStamp: (new Date()).toISOString(),
        clientVersion: 1
    };

    screw.mmd.getTermSets(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                console.log('Response:', result['soap:Envelope']['soap:Body'][0]['GetTermSetsResponse'][0]);
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getTermSets();

// ================================

var getChildTermsInTermSet = function() {

    var data = {
        baseUrl: context.siteUrl,
        sspId: config.mmd.sspId,
        termSetId: config.mmd.termSetId,
        lcid: config.mmd.lcid
    };

    screw.mmd.getChildTermsInTermSet(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                var xmlResult = result['soap:Envelope']['soap:Body'][0]['GetChildTermsInTermSetResponse'][0]['GetChildTermsInTermSetResult'][0];
                parseString(xmlResult, function(err, terms) {
                    console.log("Terms:", terms['TermStore']);
                });
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getChildTermsInTermSet();

// ================================

var getChildTermsInTerm = function() {

    var data = {
        baseUrl: context.siteUrl,
        sspId: config.mmd.sspId,
        lcid: config.mmd.lcid,
        termId: config.mmd.termId,
        termSetId: config.mmd.termSetId
    };

    screw.mmd.getChildTermsInTerm(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                var xmlResult = result['soap:Envelope']['soap:Body'][0]['GetChildTermsInTermResponse'][0]['GetChildTermsInTermResult'][0];
                parseString(xmlResult, function(err, terms) {
                    console.log("Terms:", terms['TermStore']);
                });
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getChildTermsInTerm();

// ================================

var getTermsByLabel = function() {

    var data = {
        baseUrl: context.siteUrl,
        label: "New label",
        lcid: config.mmd.lcid,
        matchOption: "ExactMatch", // StartsWith
        resultCollectionSize: 25,
        termId: config.mmd.termId,
        // termIds: [config.mmd.termId],
        addIfNotFound: false
    };

    screw.mmd.getTermsByLabel(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                var xmlResult = result['soap:Envelope']['soap:Body'][0]['GetTermsByLabelResponse'][0]['GetTermsByLabelResult'][0];
                parseString(xmlResult, function(err, terms) {
                    console.log("Terms:", terms['TermStore']);
                });
                // console.log(xmlResult);
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getTermsByLabel();

// ================================

var getKeywordTermsByGuids = function() {

    var data = {
        baseUrl: context.siteUrl,
        lcid: config.mmd.lcid,
        termId: config.mmd.termId,
        // termIds: [config.mmd.termId]
    };

    screw.mmd.getKeywordTermsByGuids(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                var xmlResult = result['soap:Envelope']['soap:Body'][0]['GetKeywordTermsByGuidsResponse'][0]['GetKeywordTermsByGuidsResult'][0];
                parseString(xmlResult, function(err, terms) {
                    console.log("Terms:", terms['TermStore']);
                });
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getKeywordTermsByGuids();

// ================================

var addTerms = function() {

    var data = {
        baseUrl: context.siteUrl,
        sspId: config.mmd.sspId,
        termSetId: config.mmd.termSetId,
        lcid: config.mmd.lcid,
        newTerms: '<newTerms>' +
                    '<newTerm label="someTerm 1" clientId="1" parentTermId="00000000-0000-0000-0000-000000000000"></newTerm>' +
                    '<newTerm label="someTerm 2" clientId="1" parentTermId="e66280de-4fdd-4cb9-8783-ce3efe3f7ef8"></newTerm>' +
                  '</newTerms>'
    };

    screw.mmd.addTerms(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                console.log("Response:", result);
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// addTerms();

// ================================

var getAllTerms = function() {

    var data = {
        baseUrl: context.siteUrl,
        serviceName: config.mmd.serviceName,
        termSetId: config.mmd.termSetId,
        properties: [
            'Id',
            'Name',
            'Description',
            'CustomProperties',
            'IsRoot',
            'IsDeprecated',
            'PathOfTerm',
            'IsAvailableForTagging',
            'Parent'
        ]
    };

    screw.mmd.getAllTerms(data)
        .then(function(response) {
            var results = JSON.parse(response.body);
            console.log("Response:", results); // [results.length - 1]['_Child_Items_']);
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getAllTerms();

// ================================

var setTermName = function() {

    var data = {
        baseUrl: context.siteUrl,
        serviceName: config.mmd.serviceName,
        termSetId: config.mmd.termSetId,
        termId: "f3b7eb21-ba15-40f1-a872-93c48f6530a2",
        newName: "New name"
    };

    screw.mmd.setTermName(data)
        .then(function(response) {
            // var results = JSON.parse(response.body);
            console.log("Response:", response.body);
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// setTermName();

// ================================

var deprecateTerm = function() {

    var data = {
        baseUrl: context.siteUrl,
        serviceName: config.mmd.serviceName,
        termSetId: config.mmd.termSetId,
        termId: "f3b7eb21-ba15-40f1-a872-93c48f6530a2",
        deprecate: true
    };

    screw.mmd.deprecateTerm(data)
        .then(function(response) {
            // var results = JSON.parse(response.body);
            console.log("Response:", response.body);
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// deprecateTerm();