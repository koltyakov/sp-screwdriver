var spf = spf || {};

spf.Screwdriver = function(context) {

    var spr = require("sp-request").create(context);

    this.ups = new (require('./api/ups'))(spr);
    this.mmd = new (require('./api/mmd'))(spr);

    return this;
};

module.exports = spf.Screwdriver;