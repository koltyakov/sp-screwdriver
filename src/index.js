const Screwdriver = function(context) {

    var spr = require("sp-request").create(context);

    this.ups = new (require('./api/ups'))(spr);
    this.mmd = new (require('./api/mmd'))(spr);
    this.versions = new (require('./api/versions'))(spr);

    return this;
};

module.exports = Screwdriver;