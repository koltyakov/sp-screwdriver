import { expect } from 'chai';
import * as path from 'path';
import { Cpass } from 'cpass';

import { JsomNode, IJsomNodeSettings } from 'sp-jsom-node';
import { Screwdriver, IScrewdriverSetting } from './../../src';
import { IEnvironmentConfig } from './../configs';

const cpass = new Cpass();

export const initScrewdriver = (testConfig: IEnvironmentConfig): Screwdriver => {
    let config: any = require(path.resolve(testConfig.configPath));
    config.password = config.password && cpass.decode(config.password);

    let screwdriverSettings: IScrewdriverSetting = {
        siteUrl: config.siteUrl,
        authOptions: config
    };

    let screw = new Screwdriver(screwdriverSettings);
    screw.init();
    return screw;
};

export const initJsom = (testConfig: IEnvironmentConfig): void => {
    let config = require(path.resolve(testConfig.configPath));
    config.password = config.password && cpass.decode(config.password);

    let jsomNodeSettings: IJsomNodeSettings = {
        siteUrl: config.siteUrl,
        authOptions: config
    };

    (new JsomNode(jsomNodeSettings)).init();
};