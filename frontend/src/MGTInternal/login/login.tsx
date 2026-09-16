import React, { useState, useRef, useEffect } from 'react';
import { msalInstance, msalLogin, msalLogout, getUserPhoto, getCachedUserProfile, getUserPresence } from './msalHelper';
import { loginRequest } from './msalConfig';
import { PresenceAvailable16Filled, PresenceBusy16Filled, PresenceDnd16Filled, PresenceAway16Filled, PresenceOffline16Regular } from '@fluentui/react-icons';
import styles from './login.module.css';
import { subscribePresence } from '../../services/graphPresenceService';
import { getNavigation } from '../../pages/Navigation';
import { NavigationItem } from '../../models/NavigationItem';
import { getToken } from '../components/getToken';

// 用户信息类型
const defaultUserDetails = {
  id: '',
  displayName: '',
  mail: '',
  personImage: '',
};

// 新增类型定义
export type UserDetails = {
  id: string;
  displayName: string;
  mail: string;
  personImage: string;
};

// presence 状态映射常量
const presenceMap: Record<string, string> = {
  available: 'Available',
  busy: 'Busy',
  away: 'Away',
  offline: 'Offline',
  dnd: 'DoNotDisturb',
  Available: 'Available',
  Busy: 'Busy',
  Away: 'Away',
  Offline: 'Offline',
  DoNotDisturb: 'DoNotDisturb',
};

export default function MgtInternalLogin({
  loginView = 'full',
  showPresence = false,
  userDetails = null,
  onLoginInitiated,
  onLoginCompleted,
  onLoginFailed,
  onLogoutInitiated,
  onLogoutCompleted,
  style = {},
}) {
  // 页面加载时自动恢复登录状态
  const [user, setUser] = useState<UserDetails | null>(() => {
    const cached = getCachedUserProfile();
    if (!cached) return null;
    return {
      id: msalInstance.getAllAccounts()[0]?.localAccountId || '',
      displayName: cached.displayName,
      mail: cached.mail,
      personImage: cached.personImage,
    };
  });
  const [flyoutOpen, setFlyoutOpen] = useState(false);
  const [presence, setPresence] = useState<'Available' | 'Busy' | 'DoNotDisturb' | 'Away' | 'Offline' | null>(null);
  const [presencePopup, setPresencePopup] = useState(false);
  const [navigationItems, setNavigationItems] = useState<NavigationItem[]>([]);
  const flyoutRef = useRef<HTMLDivElement>(null);

  // 页面加载时只恢复登录状态和presence，并补充头像
  useEffect(() => {
    async function fetchPresenceAndPhoto() {
      const cached = getCachedUserProfile();
      if (cached) {
        const rawPresence = await getUserPresence();
        setPresence((presenceMap[rawPresence ?? 'Offline'] ?? 'Offline') as typeof presence);
        const photoUrl = await getUserPhoto();
        setUser({
          id: msalInstance.getAllAccounts()[0]?.localAccountId || '',
          displayName: cached.displayName,
          mail: cached.mail,
          personImage: photoUrl || cached.personImage,
        });
        // 页面重新加载时自动订阅 presence
        const userId = msalInstance.getAllAccounts()[0]?.localAccountId || '';
        if (userId) {
          subscribePresence(userId);
        }
      }
    }
    fetchPresenceAndPhoto();
  }, []);

  // 用户状态变化时获取导航数据
  useEffect(() => {
    const fetchNavigation = async () => {
      if (user) {
        const items = await getNavigation(true);
        setNavigationItems(items);
      } else {
        setNavigationItems([]);
      }
    };
    fetchNavigation();
  }, [user]);

  // 点击弹窗外部收起弹窗
  useEffect(() => {
    if (!flyoutOpen) return;
    const handleClickOutside = (event: MouseEvent) => {
      if (flyoutRef.current && !flyoutRef.current.contains(event.target as Node)) {
        setFlyoutOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [flyoutOpen]);

  // 登录流程（微软账号，带头像和presence）
  const handleLogin = async () => {
    onLoginInitiated && onLoginInitiated();
    try {
      const account = await msalLogin();
      const profile = {
        displayName: account.name || '',
        mail: account.username || '',
        personImage: '',
      };

      setUser({
        id: account.localAccountId,
        ...profile,
      });

      if (account.localAccountId) {
        subscribePresence(account.localAccountId);
      }

      const [photoUrl, userPresence] = await Promise.all([
        getUserPhoto().catch((error) => {
          console.error('获取用户头像失败:', error);
          return null;
        }),
        getUserPresence().catch((error) => {
          console.error('获取用户状态失败:', error);
          return null;
        }),
      ]);

      profile.personImage = photoUrl || '';
      setUser({
        id: account.localAccountId,
        ...profile,
      });
      setPresence((presenceMap[userPresence ?? 'Offline'] ?? 'Offline') as typeof presence);
      onLoginCompleted && onLoginCompleted(profile);
    } catch (e) {
      console.error('Microsoft 登录失败:', e);
      onLoginFailed && onLoginFailed(e);
    }
  };

  // 登出流程（微软账号）
  const handleLogout = async () => {
    onLogoutInitiated && onLogoutInitiated();
    await msalLogout();
    setUser(null);
    setFlyoutOpen(false);
    onLogoutCompleted && onLogoutCompleted();
  };

  // 切换弹窗
  const toggleFlyout = () => setFlyoutOpen((v) => !v);

  // 状态颜色
  const getPresenceColor = () => {
    switch (presence) {
      case 'Available': return '#4caf50'; // 绿色
      case 'Busy': return '#f44336'; // 红色
      case 'Away': return '#ff9800'; // 橙色
      case 'Offline': return '#ccc'; // 灰色
      case 'DoNotDisturb': return '#f44336'; // 红色
      default: return '#ccc';
    }
  };

  // 状态选项（使用Graph API标准值）
  const presenceOptions = [
    { value: 'Available', activity: 'Available', label: 'Available', icon: 'available' },
    { value: 'Busy', activity: 'InACall', label: 'Busy', icon: 'busy' },
    { value: 'Busy', activity: 'InAConferenceCall', label: 'Do not Disturb', icon: 'dnd' },
    { value: 'Away', activity: 'Away', label: 'Appear away', icon: 'away' },
    { value: 'DoNotDisturb', activity: 'Presenting', label: 'Appear Offline', icon: 'offline' },
  ];

  // 获取accessToken
  async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) return '';
    return await getToken(msalInstance, accounts[0], loginRequest.scopes);
  }

  // 切换用户状态实际接口
  async function setPresenceApi(option: typeof presenceOptions[0]) {
    const accessToken = await getAccessToken();
    if (!accessToken || !user) return;
    const sessionId = process.env.REACT_APP_CLIENT_ID;
    if (!sessionId) {
      console.error('REACT_APP_CLIENT_ID is not configured');
      return;
    }
    // 获取用户 id（假设 userDetails 或 profile 有 id 字段）
    const userId = (user as any).id || (userDetails as any)?.id;
    if (!userId) return;
    await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/presence/setPresence`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        sessionId,
        availability: option.value,
        activity: option.activity,
        expirationDuration: 'PT1H',
      }),
    });
  }

  // 选择状态
  const handlePresenceSelect = async (option: typeof presenceOptions[0]) => {
    setPresence(option.value as any);
    setPresencePopup(false);
    await setPresenceApi(option);
  };

  // 切换状态弹窗
  const handlePresenceArrowClick = (e: React.MouseEvent) => {
    e.stopPropagation();
    setPresencePopup((v) => !v);
  };

  // 菜单图标渲染（放在组件内部）
  function renderPresenceIcon(option: typeof presenceOptions[0]) {
    switch (option.icon) {
      case 'available':
        return <PresenceAvailable16Filled className={styles.available} />;
      case 'busy':
        return <PresenceBusy16Filled className={styles.busy} />;
      case 'dnd':
        return <PresenceDnd16Filled className={styles.busy} />;
      case 'away':
        return <PresenceAway16Filled className={styles.away} />;
      case 'offline':
        return <PresenceOffline16Regular className={styles.offline} />;
      default:
        return null;
    }
  }

  // 渲染登录按钮内容
  const renderSignedOutButtonContent = () => (
    <span>Sign In</span>
  );

  // 渲染已登录按钮内容
  const renderSignedInButtonContent = () => (
    <div className={styles.accountContent}>
      <div className={styles.avatar}>
        <img
          src={user?.personImage}
          alt="avatar"
          className={styles.avatarImage}
        />
        {/* 状态图标（Fluent UI） */}
        <span className={styles.presenceBadge}>
          {renderPresenceIcon(
            presenceOptions.find(opt => opt.value === presence) || presenceOptions[0]
          )}
        </span>
      </div>
      {loginView !== 'avatar' && (
        <div className={styles.accountText}>
          <div className={styles.displayName}>{user?.displayName}</div>
          {loginView === 'full' && <div className={styles.email}>{user?.mail}</div>}
        </div>
      )}
    </div>
  );

  // 渲染弹窗
  const renderFlyout = () => (
    flyoutOpen && (
      <div
        ref={flyoutRef}
        style={{
          position: 'absolute',
          top: '100%',
          right: 0,
          background: '#fff',
          boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
          borderRadius: 12,
          padding: 24,
          zIndex: 100,
          minWidth: 320,
          maxWidth: 340,
        }}
      >
        <button
          onClick={handleLogout}
          style={{
            position: 'absolute',
            top: 16,
            right: 16,
            background: 'none',
            border: 'none',
            color: '#1976d2',
            fontWeight: 500,
            fontSize: 16,
            cursor: 'pointer',
          }}
        >
          Sign Out
        </button>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginTop: 24 }}>
          <div style={{ position: 'relative', width: 56, height: 56 }}>
            <img
              src={user?.personImage}
              alt="avatar"
              style={{ width: '100%', height: '100%', borderRadius: '50%' }}
            />
            {/* 状态图标（Fluent UI） */}
            <span style={{ position: 'absolute', right: 2, bottom: 2 }}>
              {renderPresenceIcon(
                presenceOptions.find(opt => opt.value === presence) || presenceOptions[0]
              )}
            </span>
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 600, fontSize: 18 }}>{user?.displayName}</div>
            <div style={{ fontSize: 14, color: '#666', marginBottom: 8 }}>{user?.mail}</div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, position: 'relative' }}>
              {/* 状态图标（Fluent UI，替换原圆点） */}
              {renderPresenceIcon(
                presenceOptions.find(opt => opt.value === presence) || presenceOptions[0]
              )}
              <span style={{ fontSize: 15, color: '#f44336', fontWeight: 500 }}>
                {presence === 'Busy' ? 'Busy' : presence === 'Available' ? 'Available' : presence === 'Away' ? 'Away' : presence === 'Offline' ? 'Appear Offline' : presence === 'DoNotDisturb' ? 'Do not Disturb' : ''}
              </span>
              <span
                style={{ marginLeft: 'auto', fontSize: 18, color: '#888', cursor: 'pointer' }}
                onClick={handlePresenceArrowClick}
              >&#8250;</span>
              {/* 状态切换弹窗 */}
              {presencePopup && (
                <div
                  style={{
                    position: 'absolute',
                    top: 32,
                    right: 0,
                    background: '#fff',
                    borderRadius: 10,
                    boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
                    padding: '8px 0',
                    minWidth: 160,
                    zIndex: 200,
                  }}
                >
                  {presenceOptions.map(option => (
                    <div
                      key={option.value}
                      onClick={() => handlePresenceSelect(option)}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: 8,
                        padding: '8px 16px',
                        cursor: 'pointer',
                        fontSize: 15,
                        color: '#222',
                        background: presence === option.value ? '#f5f5f5' : 'transparent',
                      }}
                    >
                      {renderPresenceIcon(option)}
                      <span>{option.label}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    )
  );

  // WebSocket 监听 presence 更新
  useEffect(() => {
    const userId = msalInstance.getAllAccounts()[0]?.localAccountId || '';
    if (!userId) return;
    const websocketUrl = `wss://${process.env.REACT_APP_DOMAIN || 'localhost:3001'}`;
    let socket: WebSocket | null = null;
    let reconnectTimer: number | undefined;
    let reconnectDelay = 1000;
    let disposed = false;

    const connect = () => {
      socket = new WebSocket(websocketUrl);
      socket.onopen = () => {
        reconnectDelay = 1000;
        console.log('Presence WebSocket connected');
      };
      socket.onmessage = async (event) => {
        const data = JSON.parse(event.data);
        if (
          data.resource === `communications/presences('${userId}')` &&
          data.changeType === 'updated'
        ) {
          const rawPresence = await getUserPresence();
          setPresence((presenceMap[rawPresence ?? 'Offline'] ?? 'Offline') as typeof presence);
        }
      };
      socket.onclose = () => {
        socket = null;
        if (!disposed) {
          reconnectTimer = window.setTimeout(connect, reconnectDelay);
          reconnectDelay = Math.min(reconnectDelay * 2, 30000);
        }
      };
      socket.onerror = (error) => {
        console.error('Presence WebSocket error:', error);
      };
    };

    connect();
    const subscriptionRenewalTimer = window.setInterval(() => {
      void subscribePresence(userId);
    }, 40 * 60 * 1000);

    return () => {
      disposed = true;
      window.clearInterval(subscriptionRenewalTimer);
      if (reconnectTimer !== undefined) {
        window.clearTimeout(reconnectTimer);
      }
      socket?.close();
    };
  }, [user]);

  return (
    <div className={styles.loginContainer} style={style}>
      {!user ? (
        <button className={styles.accountButton} onClick={handleLogin}>
          {renderSignedOutButtonContent()}
        </button>
      ) : (
        <button
          className={styles.accountButton}
          onClick={toggleFlyout}
          aria-expanded={flyoutOpen}
          aria-haspopup="menu"
        >
          {renderSignedInButtonContent()}
        </button>
      )}
      {user && renderFlyout()}
    </div>
  );
}
