const Util = function() {

    this.trimMultiline = (multiline) => {
        return multiline.split('\n').map(line => line.trim()).join('');
    };

    return this;

};

const util = new Util();
module.exports = util;