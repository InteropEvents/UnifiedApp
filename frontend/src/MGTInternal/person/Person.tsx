import React, { useEffect, useState, useRef } from 'react';
import { getToken } from '../components/getToken';
import { PersonCard } from '../personcard';
import { PersonaPresence } from '@fluentui/react';
import {
  PresenceAvailable16Filled,
  PresenceBusy16Filled,
  PresenceDnd16Filled,
  PresenceAway16Filled,
  PresenceOffline16Regular
} from '@fluentui/react-icons';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig } from '../login/msalConfig';

export type PersonCardInteraction = 'none' | 'hover' | 'click';
export type AvatarType = 'photo' | 'initials';
export type ViewType = 'image' | 'oneline' | 'twolines' | 'threelines' | 'fourlines';
export type AvatarSize = 'small' | 'large' | 'auto' | number;

export interface IDynamicPerson {
  id?: string;
  displayName?: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  userPrincipalName?: string;
  personImage?: string;
  [key: string]: any;
}

export interface PersonProps {
  personQuery?: string;
  userId?: string;
  personDetails?: IDynamicPerson;
  fallbackDetails?: IDynamicPerson;
  fetchImage?: boolean;
  disableImageFetch?: boolean;
  showPresence?: boolean;
  presence?: 'Available' | 'Busy' | 'Away' | 'Offline';
  view?: ViewType;
  verticalLayout?: boolean;
  avatarType?: AvatarType;
  avatarUrl?: string;
  avatarSize?: AvatarSize;
  line1Property?: string;
  line2Property?: string;
  line3Property?: string;
  line4Property?: string;
  style?: React.CSSProperties;
  personCardInteraction?: PersonCardInteraction;
  onUpdated?: () => void;
  onLine1Clicked?: (person: IDynamicPerson) => void;
  onLine2Clicked?: (person: IDynamicPerson) => void;
  onLine3Clicked?: (person: IDynamicPerson) => void;
  onLine4Clicked?: (person: IDynamicPerson) => void;
}

const presenceTextMap: Record<PersonaPresence, string> = {
  [PersonaPresence.online]: 'Online',
  [PersonaPresence.away]: 'Away',
  [PersonaPresence.busy]: 'Busy',
  [PersonaPresence.dnd]: 'Do Not Disturb',
  [PersonaPresence.offline]: 'Offline',
  [PersonaPresence.none]: 'Unknown',
  [PersonaPresence.blocked]: 'Unknown', // blocked
};

const presenceIconMap: Record<PersonaPresence, JSX.Element> = {
  [PersonaPresence.online]: <PresenceAvailable16Filled style={{ color: '#4caf50' }} />,
  [PersonaPresence.away]: <PresenceAway16Filled style={{ color: '#ff9800' }} />,
  [PersonaPresence.busy]: <PresenceBusy16Filled style={{ color: '#f44336' }} />,
  [PersonaPresence.dnd]: <PresenceDnd16Filled style={{ color: '#f44336' }} />,
  [PersonaPresence.offline]: <PresenceOffline16Regular style={{ color: '#ccc' }} />,
  [PersonaPresence.none]: <PresenceOffline16Regular style={{ color: '#e0e0e0' }} />,
  [PersonaPresence.blocked]: <PresenceOffline16Regular style={{ color: '#e0e0e0' }} />,
};

const presenceColor = {
  Available: '#4caf50',
  Busy: '#f44336',
  Away: '#ff9800',
  Offline: '#9e9e9e',
};

function getInitials(name?: string, email?: string) {
  if (name && name.trim()) {
    const parts = name.trim().split(' ');
    if (parts.length === 1) return parts[0][0].toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
  if (email && email[0]) return email[0].toUpperCase();
  return '?';
}

const msalInstance = new PublicClientApplication(msalConfig);
const account = msalInstance.getAllAccounts()[0] || null;

// 你需要实现此函数，返回有效的 accessToken
async function getAccessToken(): Promise<string> {
  // 使用统一的 getToken 方法获取 accessToken
  return await getToken(msalInstance, account, ['User.Read']);
}

async function fetchPerson(query: string): Promise<IDynamicPerson | null> {
  const token = await getAccessToken();
  let url = '';
  if (query.includes('@')) {
    url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(query)}`;
  } else {
    url = `https://graph.microsoft.com/v1.0/users/${query}`;
  }
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) return null;
  const data = await res.json();
  return data;
}

async function fetchPersonImage(person: IDynamicPerson): Promise<string> {
  const token = await getAccessToken();
  if (!person || !person.id) return '';
  const url = `https://graph.microsoft.com/v1.0/users/${person.id}/photo/$value`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) return '';
  const blob = await res.blob();
  return URL.createObjectURL(blob);
}

async function fetchPersonPresence(person: IDynamicPerson): Promise<'Available' | 'Busy' | 'Away' | 'Offline'> {
  const token = await getAccessToken();
  if (!person || !person.id) return 'Offline';
  const url = `https://graph.microsoft.com/v1.0/users/${person.id}/presence`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) return 'Offline';
  const data = await res.json();
  // Graph presence.activity: Available/Busy/Away/Offline 等
  if (data && data.availability) {
    if (['Available', 'Busy', 'Away', 'Offline'].includes(data.availability)) {
      return data.availability;
    }
  }
  return 'Offline';
}

// presence 字符串到 PersonaPresence 枚举的映射
const presenceStringToEnum: Record<string, PersonaPresence> = {
  available: PersonaPresence.online,
  online: PersonaPresence.online,
  busy: PersonaPresence.busy,
  dnd: PersonaPresence.dnd,
  donotdisturb: PersonaPresence.dnd,
  away: PersonaPresence.away,
  brb: PersonaPresence.away,
  offline: PersonaPresence.offline,
  none: PersonaPresence.none,
  blocked: PersonaPresence.blocked,
  unknown: PersonaPresence.none,
};

export const Person: React.FC<PersonProps> = ({
  personQuery,
  userId,
  personDetails,
  fallbackDetails,
  fetchImage = true,
  disableImageFetch = false,
  showPresence = false,
  presence,
  view = 'image',
  verticalLayout = false,
  avatarType = 'photo',
  avatarUrl,
  avatarSize = 48,
  line1Property,
  line2Property,
  line3Property,
  line4Property,
  style = {},
  personCardInteraction = 'none',
  onUpdated,
  onLine1Clicked,
  onLine2Clicked,
  onLine3Clicked,
  onLine4Clicked,
}) => {
  const [person, setPerson] = useState<IDynamicPerson | null>(personDetails || null);
  const [image, setImage] = useState<string | undefined>(avatarUrl);
  const [personPresence, setPersonPresence] = useState<string | undefined>(presence);
  const [showCard, setShowCard] = useState(false);
  const avatarRef = useRef<HTMLDivElement>(null);
  const [cardPos, setCardPos] = useState<{ left: number; top: number } | null>(null);

  // 自动查找人员信息
  useEffect(() => {
    let cancelled = false;
    async function loadPerson() {
      let p = personDetails;
      if (!p && (personQuery || userId)) {
        const fetched = await fetchPerson(personQuery || userId!);
        p = fetched === null ? undefined : fetched;
      }
      if (!p && fallbackDetails) {
        p = fallbackDetails;
      }
      if (!cancelled) setPerson(p ?? null);
      if (onUpdated) onUpdated();
    }
    loadPerson();
    return () => { cancelled = true; };
  }, [personQuery, userId, personDetails, fallbackDetails, onUpdated]);

  // 自动查找头像
  useEffect(() => {
    let cancelled = false;
    async function loadImage() {
      if (disableImageFetch) return;
      if (avatarUrl) { setImage(avatarUrl); return; }
      if (person && fetchImage) {
        const img = await fetchPersonImage(person);
        if (!cancelled) setImage(img);
      }
    }
    loadImage();
    // eslint-disable-next-line
  }, [person, avatarUrl, fetchImage, disableImageFetch]);

  // 自动查找presence
  useEffect(() => {
    let cancelled = false;
    async function loadPresence() {
      if (presence) { setPersonPresence(presence); return; }
      if (person && showPresence) {
        const p = await fetchPersonPresence(person);
        if (!cancelled) setPersonPresence(p);
      }
    }
    loadPresence();
    // eslint-disable-next-line
  }, [person, presence, showPresence]);

  // 处理行点击事件
  const handleLineClick = (idx: number) => {
    if (!person) return;
    if (idx === 0 && onLine1Clicked) onLine1Clicked(person);
    if (idx === 1 && onLine2Clicked) onLine2Clicked(person);
    if (idx === 2 && onLine3Clicked) onLine3Clicked(person);
    if (idx === 3 && onLine4Clicked) onLine4Clicked(person);
  };

  // 取各行内容
  const lines = [
    line1Property ? person?.[line1Property] : person?.displayName,
    line2Property ? person?.[line2Property] : person?.jobTitle,
    line3Property ? person?.[line3Property] : person?.department,
    line4Property ? person?.[line4Property] : person?.mail,
  ].filter((v): v is string => Boolean(v && v.trim()));

  // 头像尺寸
  const avatarPx = (typeof avatarSize === 'number' ? avatarSize : (avatarSize === 'small' ? 32 : avatarSize === 'large' ? 64 : 48)) / 2;

  // 渲染头像
  const renderAvatar = () => {
    if (avatarType === 'photo' && image) {
      return (
        <img
          src={image}
          alt={person?.displayName || person?.mail || 'person'}
          style={{ width: avatarPx, height: avatarPx, borderRadius: '50%' }}
        />
      );
    }
    return (
      <div
        style={{
          width: avatarPx,
          height: avatarPx,
          borderRadius: '50%',
          background: '#e0e0e0',
          color: '#555',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          fontSize: avatarPx / 2.4,
          fontWeight: 600,
        }}
      >
        {getInitials(person?.displayName, person?.mail)}
      </div>
    );
  };

  // 渲染 presence
  const renderPresence = () => {
    if (!showPresence || !personPresence) return null;
    // 统一转小写去空格
    const key = String(personPresence).replace(/\s+/g, '').toLowerCase();
    const enumValue = presenceStringToEnum[key] ?? PersonaPresence.none;
    return (
      <span
        style={{
          display: 'inline-block',
          position: 'absolute',
          bottom: 5,
          right: 0,
          width: 10,
          height: 10,
          zIndex: 2,
          background: 'transparent',
        }}
        title={personPresence}
      >
        {presenceIconMap[enumValue]}
      </span>
    );
  };

  // 监听 showCard 变化，计算弹窗位置
  useEffect(() => {
    if (showCard && avatarRef.current) {
      const rect = avatarRef.current.getBoundingClientRect();
      setCardPos({ left: rect.left, top: rect.bottom + 4 });
    } else {
      setCardPos(null);
    }
  }, [showCard]);

  return (
    <div
      style={{
        display: 'flex',
        flexDirection: verticalLayout ? 'column' : 'row',
        alignItems: 'center',
        gap: 12,
        position: 'relative',
        ...style,
      }}
    >
      <div
        style={{ position: 'relative', display: 'inline-block' }}
        onMouseEnter={() => setShowCard(true)}
        onMouseLeave={() => setShowCard(false)}
      >
        <div ref={avatarRef} style={{ display: 'inline-block', position: 'relative' }}>
          {renderAvatar()}
          {renderPresence()}
        </div>
        {showCard && cardPos && (
          <div
            style={{
              position: 'fixed',
              left: cardPos.left,
              top: cardPos.top-5,
              zIndex: 2000,
            }}
          >
            <PersonCard
              data={person ? {
                id: person.id,
                name: person.displayName,
                photo: image,
                jobTitle: person.jobTitle,
                department: person.department,
                mail: person.mail,
                userPrincipalName: person.userPrincipalName,
                // 可补充更多字段
              } : undefined}
            />
          </div>
        )}
      </div>
      {view !== 'image' && lines.length > 0 && (
        <div style={{ minWidth: 120 }}>
          {lines.slice(0, {
            oneline: 1,
            twolines: 2,
            threelines: 3,
            fourlines: 4,
          }[view] || 1).map((line, idx) => (
            <div
              key={idx}
              style={{ fontSize: idx === 0 ? 16 : 14, fontWeight: idx === 0 ? 600 : 400, cursor: 'pointer' }}
              onClick={() => handleLineClick(idx)}
            >
              {line}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default Person;
