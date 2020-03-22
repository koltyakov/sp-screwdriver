import { Screwdriver } from '../../src';
import { Environments as TestsConfigs } from '../configs';
import { initScrewdriver, initJsom, getTestConfigs } from './../helper';
import { ITestConfig } from '../interfaces';

declare const global: any;

for (const testConfig of TestsConfigs) {

  describe(`Run 'MMD' tests in ${testConfig.environmentName}`, () => {

    let screw: Screwdriver;
    let params: ITestConfig;
    let _spPageContextInfo: any;

    before('Configure Screwdriver', function (done: any): void {
      this.timeout(30 * 1000);
      screw = initScrewdriver(testConfig);
      initJsom(testConfig);
      _spPageContextInfo = global.window._spPageContextInfo;
      getTestConfigs(testConfig).then((testParams) => {
        params = testParams;
        done();
      }).catch(done);
    });

    it(`should get termsets via SOAP`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.getTermSets({
        sspId: params.mmd.sspId,
        termSetId: params.mmd.termSetId,
        lcid: params.mmd.lcid || 1033
      }).then((result) => {
        done();
      }).catch(done);
    });

    it(`should get get child terms in termset via SOAP`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.getChildTermsInTermSet({
        sspId: params.mmd.sspId,
        termSetId: params.mmd.termSetId,
        lcid: params.mmd.lcid || 1033
      }).then((result) => {
        done();
      }).catch(done);
    });

    it(`should get get child terms in term via SOAP`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.getChildTermsInTerm({
        sspId: params.mmd.sspId,
        termSetId: params.mmd.termSetId,
        termId: params.mmd.termId,
        lcid: params.mmd.lcid || 1033
      }).then((result) => {
        done();
      }).catch(done);
    });

    it(`should get get term by label via SOAP`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.getTermsByLabel({
        label: 'Dep1',
        termIds: [],
        lcid: params.mmd.lcid || 1033
      }).then((result) => {
        done();
      }).catch(done);
    });

    it(`should get get keyword terms by guids via SOAP`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.getKeywordTermsByGuids({
        termIds: [params.mmd.termId],
        lcid: params.mmd.lcid || 1033
      }).then((result) => {
        done();
      }).catch(done);
    });

    it(`should add term via SOAP`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.addTerms({
        sspId: params.mmd.sspId,
        termSetId: params.mmd.termSetId,
        newTerms: params.mmd.newTerms,
        lcid: params.mmd.lcid || 1033
      }).then((result) => {
        done();
      }).catch(done);
    });

    it(`should get all terms via CSOM`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.getAllTerms({
        termSetId: params.mmd.termSetId,
        serviceName: params.mmd.serviceName
      }).then((result) => {
        done();
      }).catch(done);
    });

    it(`should set term name via CSOM`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.setTermName({
        serviceName: params.mmd.serviceName,
        termSetId: params.mmd.termSetId,
        termId: params.mmd.termId,
        newName: 'New name'
      }).then((result) => {
        done();
      }).catch(done);
    });

    it(`should deprecate term via CSOM`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.mmd.deprecateTerm({
        serviceName: params.mmd.serviceName,
        termSetId: params.mmd.termSetId,
        termId: params.mmd.termId,
        deprecate: true
      }).then((result) => {
        done();
      }).catch(done);
    });

  });

}
