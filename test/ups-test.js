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

var getUserProfileByName = function() {

    var data = {
        baseUrl: context.siteUrl,
        accountName: config.data.accountName
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

var modifyUserPropertyByAccountName = function() {

    var data = {
        baseUrl: context.siteUrl,
        accountName: config.data.accountName,
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

var getUserPropertyByAccountName = function() {

    var data = {
        baseUrl: context.siteUrl,
        accountName: config.data.accountName,
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

var getUserProfilePropertyFor = function() {

    var data = {
        baseUrl: context.siteUrl,
        accountName: config.data.accountName,
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

var getPropertiesFor = function() {

    var data = {
        baseUrl: context.siteUrl,
        accountName: config.data.accountName
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