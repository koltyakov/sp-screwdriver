import { Screwdriver } from '../../src';
import { Environments as TestsConfigs } from '../configs';
import { initScrewdriver, initJsom, getTestConfigs } from './../helper';
import { ITestConfig } from '../interfaces';


declare let global: any;

for (let testConfig of TestsConfigs) {

    describe(`Run 'Authentication' test in ${testConfig.environmentName}`, () => {

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

        it(`should auth to SharePoint`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            const ctx = new SP.ClientContext(_spPageContextInfo.webServerRelativeUrl);
            const oWeb = ctx.get_web();
            ctx.load(oWeb);
            ctx.executeQueryAsync(() => {
                done();
            }, (sender, args) => {
                done(args.get_message());
            });
        });

    });

}
