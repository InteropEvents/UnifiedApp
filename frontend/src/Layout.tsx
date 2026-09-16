import React from 'react';
import { HashRouter, Route, Switch } from 'react-router-dom';
import { Header } from './components/Header';
import { SideNavigation } from './components/SideNavigation';
import { HomePage } from './pages/HomePage';
import { useIsSignedIn } from './hooks/useIsSignedIn';
import { NavigationItem } from './models/NavigationItem';
import { getNavigation } from './pages/Navigation';
import { FluentProvider, mergeClasses } from '@fluentui/react-components';
import { useAppContext } from './AppContext';
import { useState, useEffect } from 'react';
import { useLayoutStyles } from './styles/Styles';
import PubSub from 'pubsub-js';

export const Layout: React.FunctionComponent = theme => {
    const styles = useLayoutStyles();
    const [navigationItems, setNavigationItems] = React.useState<NavigationItem[]>([]);
    const [isSignedIn] = useIsSignedIn();
    const appContext = useAppContext();
    const [getHandleRemoveAPI, setHandleRemoveAPI] = useState(false);
    const [getAPIcontent, setAPIcontent] = useState<any[]>([]);
    
    const handleRemoveAPI = () => {
        if (getAPIcontent.length >= 0) {
            setHandleRemoveAPI(false); 
            setAPIcontent([]);
        }
    };
    
    const handleClearAPI = () => {
        setAPIcontent([]);
        PubSub.publish("ClearAPIdata", []);
    };

    React.useEffect(() => {
        const fetchNavigation = async () => {
            const items = await getNavigation(isSignedIn);
            setNavigationItems(items);
        };
        fetchNavigation();
    }, [isSignedIn]);
    
    React.useEffect(() => {
        const subscriptionToken = PubSub.subscribe('Calendar', async (topic, data) => {
            setAPIcontent([...data, ...getAPIcontent]);
        });
        return () => {
            PubSub.unsubscribe(subscriptionToken);
        };
    });

    React.useEffect(() => {
        // applyTheme(appContext.state.theme.key as any);
    }, [appContext]);
    
    return (
        <FluentProvider theme={appContext.state.theme.fluentTheme}>
            <div className={styles.page}>
                <HashRouter>
                    <div style={{ position: 'relative', zIndex: 1 }}>                   
                        <Header setHandleRemoveAPI={setHandleRemoveAPI} />
                    </div>
                    <div className={styles.main}>
                        <div
                            className={mergeClasses(
                                styles.sidebar,
                                `${appContext.state.sidebar.isMinimized ? styles.minimized : ''}`
                            )}
                        >
                            <SideNavigation items={navigationItems}></SideNavigation>
                        </div>
                        <div className={styles.content}>
                            <Switch>
                                {navigationItems.map(
                                    item =>
                                        ((item.requiresLogin && isSignedIn) || !item.requiresLogin) && (
                                            <Route exact={item.exact} path={item.url} children={item.component} key={item.key} />
                                        )
                                )}
                                <Route path="*" component={HomePage} />
                            </Switch>
                        </div>
                        {getHandleRemoveAPI && (
                            <div style={{ width: "800px", lineHeight: "30px", height: "100%", border: "1px solid  #ccc", padding: "5px", overflow: "auto" }}>
                                <button onClick={() => { handleClearAPI() }} style={{ fontSize: '15px', color: 'black', width: "80px", height: "20px", border: "none", textAlign: "center", backgroundColor: "#dadada", borderRadius: "24px" }}>Clear</button>
                                <p></p>
                                {getAPIcontent.map((tag, index) => (
                                    <div key={index}>
                                        {tag.type === 'GET' || tag.type === 'POST' ? (
                                            <div style={{ borderBottom: "2px solid  #ccc", paddingBottom: "20px" }}>
                                                <span><b>{tag.type}</b></span>
                                                <p style={{ margin: "0px", wordBreak: "break-all" }}><b>api:</b>{tag.api}</p>
                                            </div>
                                        ) : ""}
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </HashRouter>
            </div>
        </FluentProvider>
    );
};