import { Providers, ProviderState } from '@microsoft/mgt';
import { getGraphSdkClient } from './graphSdkClient';

export const updatePresence = async (availability: string, activity: string) => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    try {
      console.log('Udpate presence to: ', availability, activity);
      const graphClient = provider.graph.client;
      await graphClient
        .api('/me/presence/setPresence')

        .post({
          sessionId: process.env.REACT_APP_CLIENT_ID!,
          availability: availability,
          activity: activity,
          expirationDuration :"PT1H"

        });
    } catch (error) {
      const err = error as any; // Type assertion
      console.error('Error setting presence:', err+" availability:"+availability+" activity:"+activity);
    }
  } else {
    console.error('User is not signed in');
  }
};

export const updatePreferredPresence = async (availability: string, activity: string) => {
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      try {
        console.log('updatePreferredPresence to ', availability, activity);
        const graphClient = provider.graph.client;
        await graphClient
          .api('/me/presence/setUserPreferredPresence')

          .post({
            sessionId: process.env.REACT_APP_CLIENT_ID!,
            availability: availability,
            activity: activity,
            expirationDuration :"PT8H"

          });
      } catch (error) {
        const err = error as any; // Type assertion
        console.error('Error setting presence:', err+" availability:"+availability+" activity:"+activity);
      }
    } else {
      console.error('User is not signed in');
    }
  };

  export const clearPreferredPresence = async () => {
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      try {
        const graphClient = provider.graph.client;
        await graphClient
          .api('/me/presence/clearUserPreferredPresence')
          .post({});
      } catch (error) {
        const err = error as any; // Type assertion
        console.error('Error clear presence:', err);
      }
    } else {
      console.error('User is not signed in');
    }
  }

  export const getPresence = async () => {
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      try {
        const graphClient = provider.graph.client;
        const presence = await graphClient
          .api('/me/presence')
          .get();
        return presence;
      } catch (error) {
        const err = error as any; // Type assertion
        console.error('Error getting presence:', err);
      }
    } else {
      console.error('User is not signed in');
    }
  }

  export const subscribePresence = async (userId: string) => {
    try {
      const client = await getGraphSdkClient();
      const notificationUrl = `https://`+ process.env.REACT_APP_DOMAIN +`/resourceNotifications`;
      const expirationDateTime = new Date(Date.now() + 55 * 60 * 1000).toISOString();
      // 检查订阅是否已存在
      console.log('77777777777777777777777777777777777777777');
      const existingSubscriptions = await client
        .api('/subscriptions')
        .get();
      const existingSubscription = existingSubscriptions.value.find(
        (sub: any) => sub.resource === `/communications/presences/${userId}`
      );
      if (existingSubscription) {
        await client
          .api(`/subscriptions/${existingSubscription.id}`)
          .patch({ expirationDateTime });
        console.log('Subscription renewed:', existingSubscription.id);
        return;
      }
      // 创建新订阅
      const subscription = {
        changeType: 'updated',
        notificationUrl: notificationUrl,
        resource: `/communications/presences/${userId}`,
        expirationDateTime,
        clientState: 'secretClientState',
      };
      console.log('Subscription payload:', subscription);
      await client
        .api('/subscriptions')
        .post(subscription);
    } catch (error) {
      const err = error as any;
      console.error('Error subscribe:', err);
    }
  }