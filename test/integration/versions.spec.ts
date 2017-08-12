import { Screwdriver } from '../../src';
import { Environments as TestsConfigs } from '../configs';
import { initScrewdriver, initJsom } from './helper';

declare let global: any;

for (let testConfig of TestsConfigs) {

    describe(`Run 'Versions' tests in ${testConfig.environmentName}`, () => {

        let screw: Screwdriver;
        let _spPageContextInfo: any;

        before('Configure Screwdriver', function(done: any): void {
            this.timeout(30 * 1000);
            screw = initScrewdriver(testConfig);
            initJsom(testConfig);
            _spPageContextInfo = global.window._spPageContextInfo;
            done();
        });

        it(`should get document versions via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.versions.getVersions
            done();
        });

        it(`should delete all document's versions via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.versions.deleteAllVersions
            done();
        });

        it(`should delete document's version via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.versions.deleteVersion
            done();
        });

        it(`should restore document's version via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.versions.restoreVersion
            done();
        });

        it(`should get item's version collection via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.versions.getVersionCollection
            done();
        });

    });

}
