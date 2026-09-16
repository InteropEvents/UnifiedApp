import { NavigationItem } from '../models/NavigationItem';
import {
  HomeRegular,
  CalendarMailRegular,
  MailRegular,
  PeopleTeam24Regular,
  ShieldPersonRegular
} from '@fluentui/react-icons';
import { CalendarPage } from '../pages/CalendarPage';
import { HomePage } from '../pages/HomePage';
import { ChannelFilesPage } from '../pages/TeamsPage';
import { OutlookPage } from '../pages/OutlookPage';
import { EntraIDPage, hasEntraIDPermission } from '../pages/EntraIDPage';
import { WebSocketClientInstance } from '../services/WebsocketClient';


export const getNavigation = async (isSignedIn: boolean) => {
  console.log('getNavigation called with isSignedIn:', isSignedIn);
  let navItems: NavigationItem[] = [];

  navItems.push({
    name: 'Home',
    url: '/',
    icon: <HomeRegular />,
    key: 'home',
    requiresLogin: false,
    component: <HomePage />,
    exact: true
  });

  if (isSignedIn) {

    // Initialize WebSocket connection
    WebSocketClientInstance.connect();

    if (process.env.REACT_APP_ENABLE_CALENDAR === 'true') {
      navItems.push({
        name: 'Calendar',
        url: '/Calendar',
        icon: <CalendarMailRegular />,
        key: 'calendar',
        requiresLogin: true,
        component: <CalendarPage />,
        exact: true
      });
    }

    if (process.env.REACT_APP_ENABLE_OUTLOOK === 'true') {
      navItems.push({
        name: 'Outlook',
        url: '/Outlook',
        icon: <MailRegular />,
        key: 'outlook',
        requiresLogin: true,
        component: <OutlookPage />,
        exact: true
      });
    }

    if (process.env.REACT_APP_ENABLE_TEAMS === 'true') {
      navItems.push({
        name: 'Teams',
        url: '/teams',
        icon: <PeopleTeam24Regular />,
        key: 'team',
        requiresLogin: true,
        component: <ChannelFilesPage />,
        exact: true
      });
    }

    if (process.env.REACT_APP_ENABLE_ENTRAID === 'true') {
      const hasPermission = await hasEntraIDPermission();
      if (hasPermission) {
        navItems.push({
          name: 'EntraID',
          url: '/entraid',
          icon: <ShieldPersonRegular />,
          key: 'entraid',
          requiresLogin: true,
          component: <EntraIDPage />,
          exact: true
        });
      }
    }
  }

  return navItems;
};
