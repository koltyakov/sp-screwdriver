var path = require('path');
var cpass = new (require('cpass'));

var Screwdriver = require(__dirname + "/../src/index");

var configPath = path.join(__dirname + "/../config/_private.conf.json");
var context = require(configPath);
if (context.password) {
    context.password = cpass.decode(context.password);
}

var screw = new Screwdriver(context);

// ===========================================

var data = {
    baseUrl: context.siteUrl,
    accountName: 'i:0#.f|membership|testprofile@contoso.com',
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

// ===========================================