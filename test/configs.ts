import { IEnvironmentConfig } from './interfaces';

export const Environments: IEnvironmentConfig[] = [
  {
    environmentName: 'SharePoint Online',
    authConfigPath: './config/integration/private.spo.json',
    paramsConfigPath: 'private.spo'
  }
  // , {
  //   environmentName: 'On-Premise 2016',
  //   configPath: './config/integration/private.2016.json',
  //   paramsConfigPath: 'private.2016'
  // }, {
  //   environmentName: 'On-Premise 2013',
  //   configPath: './config/integration/private.2013.json',
  //   paramsConfigPath: 'private.2013'
  // }
];
