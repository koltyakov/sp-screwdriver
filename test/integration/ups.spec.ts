import { Screwdriver } from '../../src';
import { Environments as TestsConfigs } from '../configs';
import { initScrewdriver, initJsom, getTestConfigs } from './../helper';
import { ITestConfig } from '../interfaces';

declare let global: any;

for (let testConfig of TestsConfigs) {

    describe(`Run 'UPS' tests in ${testConfig.environmentName}`, () => {

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

        it(`should get user profile by name via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            screw.ups.getUserProfileByName({
                accountName: params.ups.accountName
            }).then(result => {
                done();
            }).catch(done);
        });

        it(`should get user property by account name via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            screw.ups.getUserPropertyByAccountName({
                accountName: params.ups.accountName,
                propertyName: 'SPS-Birthday'
            }).then(result => {
                done();
            }).catch(done);
        });

        it(`should modify user property by account name via SOAP`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            screw.ups.modifyUserPropertyByAccountName({
                accountName: params.ups.accountName,
                newData: [{
                    isPrivacyChanged: false,
                    isValueChanged: true,
                    privacy: 'NotSet',
                    name: 'SPS-Birthday',
                    values: [ '10.03' ]
                }]
            }).then(result => {
                done();
            }).catch(done);
        });

        it(`should get properties for via REST`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            screw.ups.getPropertiesFor({
                accountName: params.ups.accountName
            }).then(result => {
                done();
            }).catch(done);
        });

        it(`should get user profile property for via REST`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            screw.ups.getUserProfilePropertyFor({
                accountName: params.ups.accountName,
                propertyName: 'SPS-Birthday'
            }).then(result => {
                done();
            }).catch(done);
        });

        it(`should set single value profile property via CSOM`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            screw.ups.setSingleValueProfileProperty({
                accountName: params.ups.accountName,
                propertyName: 'SPS-Birthday',
                propertyValue: '10.03.1983'
            }).then(result => {
                done();
            }).catch(done);
        });

        it(`should set multi valued profile property via CSOM`, function(done: MochaDone): void {
            this.timeout(30 * 1000);
            screw.ups.setMultiValuedProfileProperty({
                accountName: params.ups.accountName,
                propertyName: 'SPS-Department',
                propertyValues: [ 'Dep1' ]
            }).then(result => {
                done();
            }).catch(done);
        });

    });

}
