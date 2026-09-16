import { Providers, ProviderState } from '@microsoft/mgt-react';
import { Client } from '@microsoft/microsoft-graph-client';

// get client instance for Microsoft Graph API
export const getGraphClient = () => {
    Providers.globalProvider.setState(ProviderState.SignedIn);
    const token = Providers.globalProvider.getAccessToken();
    const options = {
        authProvider: done => {
            done(null, token);
        }
    };
    return Client.init(options);
}