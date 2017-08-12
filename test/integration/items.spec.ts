import { Screwdriver } from '../../src';
import { Environments as TestsConfigs } from '../configs';
import { initScrewdriver, initJsom, getTestConfigs } from './../helper';
import { ITestConfig } from '../interfaces';


declare let global: any;

for (let testConfig of TestsConfigs) {

    describe(`Run 'Items' tests in ${testConfig.environmentName}`, () => {

        let screw: Screwdriver;
        let params: ITestConfig;
        let _spPageContextInfo: any;

        before('Configure Screwdriver', function(done: any): void {
            this.timeout(30 * 1000);
            screw = initScrewdriver(testConfig);
            initJsom(testConfig);
            _spPageContextInfo = global.window._spPageContextInfo;
            getTestConfigs(testConfig).then(testParams => {
                params = testParams;
                done();
            }).catch(done);
        });

        it(`should set item's property bags via CSOM`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            screw.items.setItemProperties({
                itemId: params.items.itemId,
                listPath: params.items.listPath,
                properties: params.items.properties
            }).then(result => {
                done();
            }).catch(done);
        });

    });

}
