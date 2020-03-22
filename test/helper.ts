import { expect } from 'chai';
import * as path from 'path';
import { Cpass } from 'cpass';

import { JsomNode, IJsomNodeContext } from 'sp-jsom-node';
import { Screwdriver, IScrewdriverSettings } from './../src';
import { IEnvironmentConfig, ITestConfig } from './interfaces';

const cpass = new Cpass();

export const initScrewdriver = (testConfig: IEnvironmentConfig): Screwdriver => {
  const config: any = require(path.resolve(testConfig.authConfigPath));
  config.password = config.password && cpass.decode(config.password);

  const screwdriverSettings: IScrewdriverSettings = {
    siteUrl: config.siteUrl,
    authOptions: config
  };

  const screw = new Screwdriver(screwdriverSettings);
  screw.init();
  return screw;
};

export const initJsom = (testConfig: IEnvironmentConfig): void => {
  const config = require(path.resolve(testConfig.authConfigPath));
  config.password = config.password && cpass.decode(config.password);

  const jsomNodeSettings: IJsomNodeContext = {
    siteUrl: config.siteUrl,
    authOptions: config
  };

  new JsomNode().init(jsomNodeSettings);
};

export const getTestConfigs = (testConfig: IEnvironmentConfig): Promise<ITestConfig> => {
  const modulePath: string = `./config/${testConfig.paramsConfigPath}`;
  return import(modulePath).then(conf => {
    return conf.testConfigs;
  });
};
