import ReactDOM from 'react-dom';
import { App } from './App';
import { mergeStyles } from '@fluentui/react';
import { initializeFileTypeIcons } from '@fluentui/react-file-type-icons';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import { Providers, LoginType } from '@microsoft/mgt-element';

initializeFileTypeIcons();

// Inject some global styles
mergeStyles({
  ':global(body,html,#root)': {
    margin: 0,
    padding: 0,
    height: '100vh',
    overflow: 'hidden'
  }
});

Providers.globalProvider = new Msal2Provider({
  clientId: process.env.REACT_APP_CLIENT_ID!,
  loginType: LoginType.Redirect,
  redirectUri: window.location.protocol + '//' + window.location.host,
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
    'Directory.AccessAsUser.All'
  ]
});

ReactDOM.render(<App />, document.getElementById('root'));
