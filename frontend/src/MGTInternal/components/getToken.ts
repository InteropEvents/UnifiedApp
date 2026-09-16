// src/MGTInternal/Providers/getToken.ts
// 封装获取 Microsoft Graph Token 的方法（不依赖 MGT，只用 MSAL）
import { PublicClientApplication, AccountInfo, AuthenticationResult } from '@azure/msal-browser';

/**
 * 获取 Microsoft Graph 访问令牌（AccessToken）
 * 推荐异步调用，自动处理 provider 未就绪的情况
 * @param msalInstance MSAL 实例
 * @param account 已登录账号
 * @param scopes 权限范围（可选，默认 ['User.Read']）
 * @returns Promise<string> accessToken
 */
export async function getToken(
  msalInstance: PublicClientApplication,
  account: AccountInfo | null,
  scopes: string[] = ['User.Read']
): Promise<string> {
  if (!msalInstance) throw new Error('No msalInstance provided');
  if (!account) throw new Error('No account found, please login first');
  const response: AuthenticationResult = await msalInstance.acquireTokenSilent({
    account,
    scopes,
  });
  return response.accessToken;
}

// getToken 方法已可直接用于主页面，参数为 msalInstance、account、scopes。
