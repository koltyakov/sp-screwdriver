import { Screwdriver } from '../../src';
import { Environments as TestsConfigs } from '../configs';
import { initScrewdriver, initJsom } from './helper';

declare let global: any;

for (let testConfig of TestsConfigs) {

    describe(`Run 'MMD' tests in ${testConfig.environmentName}`, () => {

        let screw: Screwdriver;
        let _spPageContextInfo: any;

        before('Configure Screwdriver', function(done: any): void {
            this.timeout(30 * 1000);
            screw = initScrewdriver(testConfig);
            initJsom(testConfig);
            _spPageContextInfo = global.window._spPageContextInfo;
            done();
        });

        it(`should get termsets via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.getTermSets
            done();
        });

        it(`should get get child terms in termset via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.getChildTermsInTermSet
            done();
        });

        it(`should get get child terms in term via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.getChildTermsInTerm
            done();
        });

        it(`should get get term by label via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.getTermsByLabel
            done();
        });

        it(`should get get keyword terms by guids via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.getKeywordTermsByGuids
            done();
        });

        it(`should add term via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.addTerms
            done();
        });

        it(`should get all terms via CSOM`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.getAllTerms
            done();
        });

        it(`should set term name via CSOM`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.setTermName
            done();
        });

        it(`should deprecate term via CSOM`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            // screw.mmd.deprecateTerm
            done();
        });

    });

}
