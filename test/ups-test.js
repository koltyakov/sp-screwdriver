const path = require('path');
const cpass = new (require('cpass'));
const parseString = require('xml2js').parseString;

const Screwdriver = require(__dirname + "/../src/index");

let configPath = path.join(__dirname + "/config/_private.conf.json");
let config = require(configPath);
let context = config.context;
if (context.password) {
    context.password = cpass.decode(context.password);
}

const screw = new Screwdriver(context);

// ================================

let getUserProfileByName = () => {

    let data = {
        baseUrl: context.siteUrl,
        accountName: config.ups.accountName
    };

    screw.ups.getUserProfileByName(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                console.log('Response:', result['soap:Envelope']['soap:Body'][0]['GetUserProfileByNameResponse'][0]['GetUserProfileByNameResult'][0]);
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getUserProfileByName();

// ================================

let modifyUserPropertyByAccountName = () => {

    let data = {
        baseUrl: context.siteUrl,
        accountName: config.ups.accountName,
        newData: [{
            isPrivacyChanged: false,
            isValueChanged: true,
            privacy: 'NotSet',
            name: 'SPS-Birthday',
            values: [ '10.03' ]
        }, {
            isPrivacyChanged: false,
            isValueChanged: true,
            privacy: 'NotSet',
            name: 'SPS-Department',
            values: [ 'Administration' ]
        }]
    };

    screw.ups.modifyUserPropertyByAccountName(data)
        .then(function(response) {
            console.log('Response:', response.body);
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// modifyUserPropertyByAccountName();

// ================================

let getUserPropertyByAccountName = () => {

    let data = {
        baseUrl: context.siteUrl,
        accountName: config.ups.accountName,
        propertyName: 'SPS-Birthday'
    };

    screw.ups.getUserPropertyByAccountName(data)
        .then(function(response) {
            parseString(response.body, function(err, result) {
                console.log('Response:', result['soap:Envelope']['soap:Body'][0]['GetUserPropertyByAccountNameResponse'][0]['GetUserPropertyByAccountNameResult'][0]);
            });
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getUserPropertyByAccountName();

// ================================

let getUserProfilePropertyFor = () => {

    let data = {
        baseUrl: context.siteUrl,
        accountName: config.ups.accountName,
        propertyName: 'SPS-Birthday'
    };

    screw.ups.getUserProfilePropertyFor(data)
        .then(function(response) {
            console.log('Response:', response.body);
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getUserProfilePropertyFor();

// ================================

let getPropertiesFor = () => {

    let data = {
        baseUrl: context.siteUrl,
        accountName: config.ups.accountName
    };

    screw.ups.getPropertiesFor(data)
        .then(function(response) {
            console.log('Response:', response.body);
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// getPropertiesFor();

// ================================

let setSingleValueProfileProperty = () => {

    let data = {
        baseUrl: context.siteUrl,
        accountName: config.ups.accountName,
        propertyName: 'AboutMe',
        propertyValue: 'Front-end developer & Lazy guy'
    };

    screw.ups.setSingleValueProfileProperty(data)
        .then(function(response) {
            console.log('Response:', JSON.parse(response.body));
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
// setSingleValueProfileProperty();

// ================================

let setMultiValuedProfileProperty = () => {

    let data = {
        baseUrl: context.siteUrl,
        accountName: config.ups.accountName,
        propertyName: 'SPS-Skills',
        propertyValues: [ 'Git', 'Node.js', 'JavaScript', 'SharePoint' ]
    };

    screw.ups.setMultiValuedProfileProperty(data)
        .then(function(response) {
            console.log('Response:', JSON.parse(response.body));
        })
        .catch(function(err) {
            console.log('Error:', err.message);
        });

};
setMultiValuedProfileProperty();