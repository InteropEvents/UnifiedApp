import * as React from 'react';
import { useIsSignedIn } from '../hooks/useIsSignedIn';
import { useHistory } from 'react-router-dom';
import { useAppContext } from '../AppContext';
import { Label, mergeClasses } from '@fluentui/react-components';
import { GridDotsRegular } from '@fluentui/react-icons';
import { useLocation } from 'react-router-dom';
import { SettingsDialog } from './SettingsDailog';
import { HeaderProps } from './Interface';
import { useHeaderStyles } from '../styles/Styles';
import MgtInternalLogin from"../MGTInternal/login/login";

const HeaderComponent: React.FunctionComponent<HeaderProps> = ({ setHandleRemoveAPI }) => {
  const styles = useHeaderStyles();
  const [isSignedIn] = useIsSignedIn();
  const appContext = useAppContext();
  const setAppContext = appContext.setState;
  const location = useLocation();
  const history = useHistory();

  const onSearchTermChanged = (e: CustomEvent) => {
    if (!(e.detail === '' && appContext.state.searchTerm === '*') && e.detail !== appContext.state.searchTerm) {
      appContext.setState({ ...appContext.state, searchTerm: e.detail === '' ? '*' : e.detail });

      if (e.detail === '') {
        history.push('/search');
      } else {
        history.push('/search?q=' + e.detail);
      }
    }
  };

  React.useLayoutEffect(() => {
    if (location.pathname === '/search') {
      const searchTerm = decodeURI(location.search.replace('?q=', ''));
      setAppContext(previous => {
        return { ...previous, searchTerm: searchTerm === '' ? '*' : searchTerm };
      });
    }
  }, [location, setAppContext]);

  return (
    <div className={styles.header} >
      <div className={styles.waffle} >
              <div className={styles.waffleLogo} >
                  <a href="https://www.office.com/apps?auth=2" target="_blank" rel="noopener noreferrer" aria-label="Open Office apps">
                     <GridDotsRegular className={styles.waffleIcon} />
                   </a>
              </div>
              {/*显示左侧的字体*/}
              <div className={styles.waffleTitle} >
                  <Label className={styles.name} >{process.env.REACT_APP_SITE_NAME} </Label>
              </div>
        
          </div>
      {/*移到最右侧*/}
          <div className={styles.login}>
          <SettingsDialog setHandleRemoveAPI={setHandleRemoveAPI} />
          <div className={mergeClasses(!isSignedIn ? styles.signedOut : styles.signedIn, styles.root)}>
                  {/* <Login>
                      <SimpleLogin template="signed-in-button-content" />
                      <LoginFlyout template="flyout-person-details"/>
                  </Login> */}
                  <MgtInternalLogin
                    loginView="full"
                    onLoginInitiated={() => {}}
                    onLoginCompleted={user => console.log('登录成功', user)}
                    onLoginFailed={error => console.log('登录失败', error)}
                    onLogoutInitiated={() => {}}
                    onLogoutCompleted={() => console.log('已登出')}
                  />
              </div>
          </div>
    </div>
  );
};

export const Header = React.memo(HeaderComponent);
