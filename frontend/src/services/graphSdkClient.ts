import { Client } from '@microsoft/microsoft-graph-client';
import { Providers, ProviderState } from '@microsoft/mgt';

export const getGraphSdkClient = async () => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const accessToken = await provider.getAccessToken();
    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
  }
  throw new Error('User is not signed in');
};
