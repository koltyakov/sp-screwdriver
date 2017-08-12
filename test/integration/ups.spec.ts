import { Screwdriver } from '../../src';
import { Environments as TestsConfigs } from '../configs';
import { initScrewdriver, initJsom } from './helper';

declare let global: any;

for (let testConfig of TestsConfigs) {

    describe(`Run 'UPS' tests in ${testConfig.environmentName}`, () => {

        let screw: Screwdriver;
        let _spPageContextInfo: any;

        before('Configure Screwdriver', function(done: any): void {
            this.timeout(30 * 1000);
            screw = initScrewdriver(testConfig);
            initJsom(testConfig);
            _spPageContextInfo = global.window._spPageContextInfo;
            done();
        });

        it(`should get user profile by name via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.ups.getUserProfileByName
            done();
        });

        it(`should get user property by account name via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.ups.getUserPropertyByAccountName
            done();
        });

        it(`should modify user property by account name via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.ups.modifyUserPropertyByAccountName
            done();
        });

        it(`should get properties for via REST`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.ups.getPropertiesFor
            done();
        });

        it(`should get user profile property for via REST`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.ups.getUserProfilePropertyFor
            done();
        });

        it(`should set single value profile property via CSOM`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.ups.setSingleValueProfileProperty
            done();
        });

        it(`should set multi valued profile property via CSOM`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.ups.setMultiValuedProfileProperty
            done();
        });

    });

}
