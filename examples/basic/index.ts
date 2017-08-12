import { Screwdriver, IScrewdriverSettings } from 'sp-screwdriver';

const settings: IScrewdriverSettings = {
    // ...
};

const screw = new Screwdriver(settings);

screw.wizard().then(() => {

    screw.ups.getPropertiesFor({
        accountName: 'i:0#.f|membership|username'
    }).then(result => {
        // ...
    }).catch(console.log);

    screw.mmd.getAllTerms({
        serviceName: 'Taxonomy_5KSgChEZ9j15+7UVInQNRQ==',
        termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f'
    }).then(result => {
        // ...
    }).catch(console.log);

});
