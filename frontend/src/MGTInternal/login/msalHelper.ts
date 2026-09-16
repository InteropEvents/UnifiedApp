// 安装依赖：npm install @azure/msal-browser
import { PublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { msalConfig, loginRequest } from './msalConfig';

export const msalInstance = new PublicClientApplication(msalConfig);

export async function msalLogin(): Promise<AccountInfo> {
  const loginResponse = await msalInstance.loginPopup(loginRequest);
  if (!loginResponse.account) {
    throw new Error('Microsoft 登录成功，但响应中没有账户信息');
  }
  return loginResponse.account;
}

export async function msalLogout() {
  await msalInstance.logoutRedirect();
}

export async function getUserProfile(): Promise<{
  displayName: string;
  mail: string;
  personImage: string;
} | null> {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;
  const account = accounts[0];
  // 这里只返回基本信息，头像需调用Graph API
  return {
    displayName: account.name || '',
    mail: account.username || '',
    personImage: '', // 头像需另行获取
  };
}

// 新增：获取用户头像
export async function getUserPhoto() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;
  const accessToken = await msalInstance.acquireTokenSilent({
    ...loginRequest,
    account: accounts[0],
  });
  const response = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
    headers: {
      Authorization: `Bearer ${accessToken.accessToken}`,
    },
  });
  if (response.ok) {
    const blob = await response.blob();
    return URL.createObjectURL(blob);
  }
  return null;
}

// 新增：获取用户在线状态
export async function getUserPresence(): Promise<'available' | 'offline' | 'busy' | 'away' | null> {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;
  const accessToken = await msalInstance.acquireTokenSilent({
    ...loginRequest,
    account: accounts[0],
  });
  const response = await fetch('https://graph.microsoft.com/v1.0/me/presence', {
    headers: {
      Authorization: `Bearer ${accessToken.accessToken}`,
    },
  });
  if (response.ok) {
    const data = await response.json();
    switch (data.availability) {
      case 'Available': return 'available';
      case 'Busy': return 'busy';
      case 'Away': return 'away';
      case 'Offline': return 'offline';
      default: return null;
    }
  }
  return null;
}

// 检查并恢复登录状态
export function getCachedUserProfile(): {
  displayName: string;
  mail: string;
  personImage: string;
} | null {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;
  const account = accounts[0];
  return {
    displayName: account.name || '',
    mail: account.username || '',
    personImage: '', // 头像需另行获取
  };
}
