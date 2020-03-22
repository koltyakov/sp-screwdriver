import { AuthConfig as SPAuthConfigirator } from 'node-sp-auth-config';
import * as colors from 'colors';
import * as path from 'path';

import { Environments } from './configs';

export async function checkOrPromptForIntegrationConfigCreds (): Promise<void> {

  const configs = [];

  for (const testConfig of Environments) {
    console.log(`\n=== ${colors.bold.yellow(`${testConfig.environmentName} Credentials`)} ===\n`);
    await (new SPAuthConfigirator({
      configPath: testConfig.authConfigPath
    })).getContext();
    console.log(colors.grey(`Gotcha ${path.resolve(testConfig.authConfigPath)}`));
  }

  console.log('\n');

}

checkOrPromptForIntegrationConfigCreds();
