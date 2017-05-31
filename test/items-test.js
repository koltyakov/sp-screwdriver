const path = require('path');
const cpass = new (require('cpass')).Cpass();
// const parseString = require('xml2js').parseString;

const Screwdriver = require(__dirname + "/../src/index");

let configPath = path.join(__dirname + "/config/private.json");
let config = require(configPath);
let context = config.context;
if (context.password) {
    context.password = cpass.decode(context.password);
}

const screw = new Screwdriver(context);

// ================================

let setItemProperties = () => {

    let data = {
        baseUrl: context.siteUrl,
        listPath: `${config.context.siteUrl}${config.items.listPath}`,
        itemId: config.items.itemId,
        properties: [{
            field: 'PublishingAssociatedContentType',
            value: ';#Article Page;#0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900242457EFB8B24247815D688C526CD44D;#'
        }]
    };

    screw.items.setItemProperties(data)
        .then(response => {
            console.log('Response:', JSON.parse(response.body)[2]);
        })
        .catch(err => {
            console.log('Error:', err.message);
        });

};
setItemProperties();