// MSAL 配置示例
export const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_CLIENT_ID!,
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: window.location.protocol + '//' + window.location.host,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: [
    'Bookmark.Read.All',
    'Calendars.ReadWrite',
    'ExternalItem.Read.All',
    'Files.Read',
    'Files.Read.All',
    'Files.ReadWrite.All',
    'Group.Read.All',
    'Group.ReadWrite.All',
    'Mail.Read',
    'Mail.ReadBasic',
    'Mail.ReadWrite',
    'OnlineMeetingTranscript.Read.All',
    'People.Read',
    'People.Read.All',
    'Presence.Read.All',
    'Presence.ReadWrite',
    'User.Read',
    'Sites.Read.All',
    'Sites.ReadWrite.All',
    'Tasks.Read',
    'Tasks.ReadWrite',
    'Team.ReadBasic.All',
    'User.ReadBasic.All',
    'User.ReadWrite.All',
    'Directory.ReadWrite.All',
    'Directory.AccessAsUser.All',
    'Channel.ReadBasic.All'
  ],
};
