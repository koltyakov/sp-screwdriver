import { Screwdriver } from '../../src';
import { Environments as TestsConfigs } from '../configs';
import { initScrewdriver, initJsom, getTestConfigs } from './../helper';
import { ITestConfig } from '../interfaces';

declare const global: any;

for (const testConfig of TestsConfigs) {

  describe(`Run 'Versions' tests in ${testConfig.environmentName}`, () => {

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

    it(`should get document versions via SOAP`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.versions.getVersions({
        fileName: params.versions.documents.fileName
      }).then((result) => {
        done();
      }).catch(done);
    });

    // Do not work?
    // it(`should delete all document's versions via SOAP`, function(done: MochaDone): void {
    //     this.timeout(30 * 1000);
    //     screw.versions.deleteAllVersions({
    //         fileName: params.versions.documents.fileName
    //     }).then(result => {
    //         done();
    //     }).catch(done);
    // });

    // Do not work?
    // it(`should delete document's version via SOAP`, function(done: MochaDone): void {
    //     this.timeout(30 * 1000);
    //     screw.versions.deleteVersion({
    //         fileName: params.versions.documents.fileName,
    //         fileVersion: '1.0'
    //     }).then(result => {
    //         done();
    //     }).catch(done);
    // });

    // Do not work?
    // it(`should restore document's version via SOAP`, function(done: MochaDone): void {
    //     this.timeout(30 * 1000);
    //     screw.versions.restoreVersion({
    //         fileName: params.versions.documents.fileName,
    //         fileVersion: '1.0'
    //     }).then(result => {
    //         done();
    //     }).catch(done);
    // });

    it(`should get item's version collection via SOAP`, function (done: MochaDone): void {
      this.timeout(30 * 1000);
      screw.versions.getVersionCollection({
        listId: params.versions.items.listId,
        itemId: params.versions.items.itemId,
        fieldName: params.versions.items.fieldName
      }).then((result) => {
        done();
      }).catch(done);
    });

  });

}
