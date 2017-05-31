const path = require('path');
const cpass = new (require('cpass')).Cpass();
const parseString = require('xml2js').parseString;

const Screwdriver = require(__dirname + "/../src/index");

let configPath = path.join(__dirname + "/config/private.json");
let config = require(configPath);
let context = config.context;
if (context.password) {
    context.password = cpass.decode(context.password);
}

const screw = new Screwdriver(context);

// ================================

let getDocumentVersions = () => {

    let data = {
        baseUrl: context.siteUrl,
        fileName: config.versions.documents.fileName
    };

    screw.versions.getVersions(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                console.log('Response:', result['soap:Envelope']['soap:Body'][0]['GetVersionsResponse'][0]['GetVersionsResult'][0]['results'][0]['result']);
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getDocumentVersions();

// ================================

let getItemVersions = () => {

    let data = {
        baseUrl: context.siteUrl,
        listId: config.versions.items.listId,
        itemId: config.versions.items.itemId,
        fieldName: config.versions.items.fieldName
    };

    screw.versions.getVersionCollection(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                console.log('Response:', result['soap:Envelope']['soap:Body'][0]['GetVersionCollectionResponse'][0]['GetVersionCollectionResult'][0]['Versions'][0]['Version']);
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
getItemVersions();