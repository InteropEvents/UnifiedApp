import React, { useEffect } from 'react';
import { Client } from '@microsoft/microsoft-graph-client';
import { PublicClientApplication } from "@azure/msal-browser";
import { getToken } from './getToken';
import { msalConfig,loginRequest } from '../login/msalConfig';

export enum ProviderState {
    SignedIn = 'SignedIn',
    SignedOut = 'SignedOut',
    Loading = 'Loading',
}

// mgt风格全局单例实现
class GlobalProvider {
    state = ProviderState.Loading;
    accessToken = '';
    graph: any = null;
    msalInstance;
    loginRequest;
    client: any = null; // 添加 client 属性


    // 新增：初始化 graph 客户端
    initGraph(token: string) {
        this.graph = Client.init({
            authProvider: (done) => {
                done(null, token);
            }
        });
    }

    async getAccessToken() {
        const msalInstance = new PublicClientApplication(msalConfig);
        const accounts = msalInstance.getAllAccounts();
            if (!accounts.length) return '';
            const token = await getToken(msalInstance, accounts[0], loginRequest.scopes);
            this.accessToken = token;
            this.initGraph(token); // 获取 token 后初始化 graph
            return token;
    }

    setState(newState) {
        console.log('[GlobalProvider] setState:', newState);
        this.state = newState;
    }
}

export const Providers = {
    globalProvider: new GlobalProvider()
};
console.log('[Providers] 导出:', Providers);

// 用法示例：
// import { Providers } from './Providers';
// const { state, getAccessToken, graph } = Providers.globalProvider;
