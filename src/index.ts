import * as path from 'path';
import * as https from 'https';

import { AuthConfig as SPAuthConfigirator } from 'node-sp-auth-config';
import { create as createRequest, ISPRequest } from 'sp-request';
import { Cpass } from 'cpass';

import { IScrewdriverSettings } from './interfaces/IScrewdriver';

import { UPS } from './api/ups';
import { MMD } from './api/mmd';
import { Versions } from './api/versions';
import { Items } from './api/items';

export class Screwdriver {

  public ups: UPS;
  public mmd: MMD;
  public versions: Versions;
  public items: Items;

  private settings: IScrewdriverSettings;
  private spAuthConfigirator: SPAuthConfigirator;
  private request: ISPRequest;
  private agent: https.Agent;

  constructor (settings: IScrewdriverSettings = {}) {
    let config = settings.config || {};
    this.settings = {
      ...settings,
      config: {
        ...config,
        configPath: path.resolve(config.configPath || path.join(process.cwd(), 'config/private.json')),
        encryptPassword: typeof config.encryptPassword !== 'undefined' ? config.encryptPassword : true,
        saveConfigOnDisk: typeof config.saveConfigOnDisk !== 'undefined' ? config.saveConfigOnDisk : true
      }
    };
    if (typeof this.settings.authOptions !== 'undefined') {
      const cpass = new Cpass();
      (this.settings.authOptions as any).password = (this.settings.authOptions as any).password &&
        cpass.decode((this.settings.authOptions as any).password);
    }
    this.agent = new https.Agent({
      rejectUnauthorized: false,
      keepAlive: true,
      keepAliveMsecs: 10000
    });
    this.spAuthConfigirator = new SPAuthConfigirator(this.settings.config);
  }

  // Init Screwdriver environment
  public init = () => {
    this.request = createRequest(this.settings.authOptions);

    this.ups = new UPS(this.request, this.settings.siteUrl);
    this.mmd = new MMD(this.request, this.settings.siteUrl);
    this.versions = new Versions(this.request, this.settings.siteUrl);
    this.items = new Items(this.request, this.settings.siteUrl);
  }

  // Trigger wizard and init Screwdriver environment
  public wizard (): Promise<IScrewdriverSettings> {
    return new Promise((resolve, reject) => {
      if (typeof this.settings.authOptions === 'undefined') {
        this.spAuthConfigirator.getContext()
          .then((context) => {
            const cpass = new Cpass();
            (context.authOptions as any).password = (context.authOptions as any).password &&
              cpass.decode((context.authOptions as any).password);
            this.settings = {
              ...this.settings,
              ...context as any
            };
            this.init();
            resolve(this.settings);
          })
          .catch((error: any) => {
            reject(error);
          });
      } else {
        resolve(this.settings);
      }
    });
  }

}

export { IScrewdriverSettings } from './interfaces/IScrewdriver';
