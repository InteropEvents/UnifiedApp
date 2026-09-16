import React, { useEffect, useState } from 'react';
import { Client } from '@microsoft/microsoft-graph-client';
import { Persona, PersonaSize, PersonaPresence } from '@fluentui/react';
import {
  PresenceAvailable16Filled,
  PresenceBusy16Filled,
  PresenceDnd16Filled,
  PresenceAway16Filled,
  PresenceOffline16Regular,
  Sparkle20Regular
} from '@fluentui/react-icons';
import { Button } from '@fluentui/react-components';
import './Agenda.css';
import { PersonCard, PersonCardProps } from '../personcard/PersonCard';
import PubSub from 'pubsub-js';
import { parseLocalDateKey } from './calendarDate';
import { getDemoMeetingTranscript } from './demoMeetingTranscript';
import {
  getAzureOpenAIChatCompletionsUrl,
  summarizeMeetingTranscript,
} from '../../services/meetingSummaryService';

type EventAttendee = {
  email: string;
  name: string;
  id: string;
  photo?: string;
  presence?: PersonaPresence;
};

type EventItem = {
  id: string;
  subject: string;
  start: string;
  end: string;
  attendees: EventAttendee[];
  organizer: EventAttendee;
  location: string;
};

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

export const Agenda: React.FC<AgendaProps> = ({
  groupByDay = false,
  eventQuery,
  id,
  agendaKey, // 替换 key
  token,
  children,
}) => {
  const [events, setEvents] = useState<EventItem[]>([]);
  const [selectedUser, setSelectedUser] = useState<PersonCardProps['data'] | null>(null);
  const [cardPosition, setCardPosition] = useState<{left: number; top: number} | null>(null);
  const [hoverOnAvatar, setHoverOnAvatar] = useState(false);
  const [hoverOnCard, setHoverOnCard] = useState(false);
  const [summaries, setSummaries] = useState<Record<string, string>>({});
  const [summaryErrors, setSummaryErrors] = useState<Record<string, string>>({});
  const [summaryLoadingEventId, setSummaryLoadingEventId] = useState<string | null>(null);

  // Graph Client
  const getGraphClient = async () => {
    const accessToken = await token();
    return Client.init({
      authProvider: done => done(null, accessToken),
    });
  };

  // 批量获取 presence
  const getPresences = async (client: any, ids: string[]): Promise<Record<string, PersonaPresence>> => {
    try {
      const apiUrl = 'https://graph.microsoft.com/v1.0/communications/getPresencesByUserId';
      const res = await client.api('/communications/getPresencesByUserId').post({ ids });
      // res.value: [{id, availability, activity}]
      const result: Record<string, PersonaPresence> = {};
      (res.value || []).forEach((item: any) => {
        const availability = item.availability || item.Availability || item.status || 'Offline';
        result[item.id] = presenceMap[availability] ?? PersonaPresence.none;
      });
      return result;
    } catch {
      // 失败时全部返回 none
      return Object.fromEntries(ids.map(id => [id, PersonaPresence.none]));
    }
  };

  // 获取日程
  useEffect(() => {
    (async () => {
      const client = await getGraphClient();
      const query = eventQuery || '/me/events';
      const apiUrl = `https://graph.microsoft.com/v1.0${query.startsWith('/') ? query : '/' + query}`;
      const req = client.api(query)
        .select('subject,start,end,attendees,organizer,location')
        .orderby('start/dateTime')
        .top(50);
      const res = await req.get();
      // 收集所有需要查 presence 的邮箱
      let allEmails: string[] = [];
      res.value.forEach((e: any) => {
        if (e.organizer?.emailAddress?.address) allEmails.push(ensureEmail(e.organizer.emailAddress.address));
        (e.attendees || []).forEach((a: any) => {
          if (a.emailAddress?.address) allEmails.push(ensureEmail(a.emailAddress.address));
        });
      });
      allEmails = Array.from(new Set(allEmails));
      // 批量查用户ID
      const emailToId: Record<string, string> = {};
      await Promise.all(allEmails.map(async email => {
        try {
          const userApi = `https://graph.microsoft.com/v1.0/users/${email}`;
          PubSub.publish('Calendar', [{ api: userApi + '?$select=id', type: 'GET' }]);
          const user = await client.api(`/users/${email}`).select('id').get();
          emailToId[email] = user.id;
        } catch {
          emailToId[email] = email; // 查不到就用邮箱本身
        }
      }));
      // 用ID批量查presence
      const allIds = Object.values(emailToId);
      const presenceApi = 'https://graph.microsoft.com/v1.0/communications/getPresencesByUserId';
      PubSub.publish('Calendar', [{ api: presenceApi, type: 'POST' }]);
      const presenceMapResult = await getPresences(client, allIds);
      const items: EventItem[] = await Promise.all(res.value.map(async (e: any) => {
        // 获取头像和 presence
        const getPhoto = async (email: string) => {
          const safeEmail = ensureEmail(email);
          try {
            const photoApi = `https://graph.microsoft.com/v1.0/users/${safeEmail}/photo/$value`;
            PubSub.publish('Calendar', [{ api: photoApi, type: 'GET' }]);
            const photoBlob = await client.api(`/users/${safeEmail}/photo/$value`).get();
            return URL.createObjectURL(photoBlob);
          } catch {
            return undefined;
          }
        };
        const attendees: EventAttendee[] = await Promise.all(
          (e.attendees || []).map(async (a: any) => {
            const safeEmail = ensureEmail(a.emailAddress.address);
            const userId = emailToId[safeEmail] || safeEmail;
            return {
              email: safeEmail,
              name: a.emailAddress.name,
              id: userId, // id 现在为用户ID或邮箱
              photo: await getPhoto(safeEmail),
              presence: presenceMapResult[userId] ?? PersonaPresence.none,
            };
          })
        );
        const organizerEmail = ensureEmail(e.organizer.emailAddress.address);
        const organizerId = emailToId[organizerEmail] || organizerEmail;
        const organizer: EventAttendee = {
          email: organizerEmail,
          name: e.organizer.emailAddress.name,
          id: organizerId,
          photo: await getPhoto(organizerEmail),
          presence: presenceMapResult[organizerId] ?? PersonaPresence.none,
        };
        return {
          id: e.id,
          subject: e.subject,
          start: e.start.dateTime,
          end: e.end.dateTime,
          attendees,
          organizer,
          location: e.location?.displayName || '',
        };
      }));
      setEvents(items);
    })();
  }, [eventQuery, agendaKey, token]); // 替换 key

  // 分组逻辑
  const groupedEvents = groupByDay
    ? events.reduce((acc, ev) => {
        // 先将 start 解析为 UTC，再转成本地日期字符串
        const startUTC = new Date(ev.start + (ev.start.endsWith('Z') ? '' : 'Z'));
        const day = startUTC.getFullYear() + '-' + String(startUTC.getMonth() + 1).padStart(2, '0') + '-' + String(startUTC.getDate()).padStart(2, '0');
        acc[day] = acc[day] || [];
        acc[day].push(ev);
        return acc;
      }, {} as Record<string, EventItem[]>)
    : null;

  // 点击头像弹出人物卡
  const handleAvatarClick = async (user: EventAttendee) => {
    // 通过 Graph API 获取更多用户信息
    let phones: string[] = [];
    let address = '';
    let mobile = '';
    try {
      const client = await getGraphClient();
      const userApi = `https://graph.microsoft.com/v1.0/users/${user.id}`;
      PubSub.publish('Calendar', [{ api: userApi + '?$select=businessPhones,mobilePhone,officeLocation', type: 'GET' }]);
      const userInfo = await client.api(`/users/${user.id}`)
        .select('businessPhones,mobilePhone,officeLocation')
        .get();
      phones = userInfo.businessPhones || [];
      mobile = userInfo.mobilePhone || '';
      address = userInfo.officeLocation || '';
    } catch (e) {
      // 获取失败时保持默认
      console.warn('获取用户详细信息失败', e);
    }
    setSelectedUser({
      id: user.id,
      name: user.name,
      photo: user.photo,
      mail: user.email,
      presenceText: user.presence !== undefined ? presenceTextMap[user.presence] : '',
      phones,
      address,
      mobile,
      // 可扩展更多字段
    });
  };

  // 鼠标移入头像时，定位弹窗
  const handleAvatarMouseEnter = (user: EventAttendee, e: React.MouseEvent<HTMLSpanElement>) => {
    setHoverOnAvatar(true);
    // 获取头像元素位置
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    // 计算弹窗位置（fixed 定位，直接用 rect）
    const cardWidth = 320; // 假设卡片宽度
    const cardHeight = 400; // 假设卡片高度
    let left = rect.left;
    let top = rect.bottom + 4;
    // 右侧溢出时左移
    if (left + cardWidth > window.innerWidth) {
      left = window.innerWidth - cardWidth - 8;
    }
    // 下方溢出时上移
    if (top + cardHeight > window.innerHeight) {
      top = rect.top - cardHeight - 4;
    }
    setCardPosition({ left, top });
    handleAvatarClick(user);
  };

  const handleAvatarMouseLeave = () => {
    setHoverOnAvatar(false);
    setTimeout(() => {
      if (!hoverOnCard) {
        setSelectedUser(null);
        setCardPosition(null);
      }
    }, 100);
  };

  const handleCardMouseEnter = () => {
    setHoverOnCard(true);
  };
  const handleCardMouseLeave = () => {
    setHoverOnCard(false);
    setTimeout(() => {
      if (!hoverOnAvatar) {
        setSelectedUser(null);
        setCardPosition(null);
      }
    }, 100);
  };

  // 渲染与会者头像，溢出时显示+N
  const renderAttendees = (attendees: EventAttendee[]) => {
    const max = 5;
    const visible = attendees.slice(0, max);
    const overflow = attendees.length - max;
    return (
      <>
        {visible.map((user, idx) => (
          <span
            key={user.id + '_' + idx}
            className="agenda-avatar"
            style={{ position: 'relative', display: 'inline-block' }}
            title={user.email}
            onMouseEnter={e => handleAvatarMouseEnter(user, e)}
            onMouseLeave={handleAvatarMouseLeave}
            tabIndex={0}
            role="button"
            aria-label={`查看${user.name}的资料`}
          >
            <Persona
              text={user.name}
              size={PersonaSize.size32}
              imageUrl={user.photo}
              hidePersonaDetails
            />
            {/* presence 图标，绝对定位右下角，样式与 login.tsx 一致 */}
            <span style={{ position: 'absolute', right: 0, bottom: 0, width: 18, height: 18, zIndex: 2, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              {presenceIconMap[user.presence ?? PersonaPresence.none]}
            </span>
          </span>
        ))}
        {overflow > 0 && (
          <span className="agenda-attendees-overflow">+{overflow}</span>
        )}
      </>
    );
  };

  const handleGenerateDemoSummary = async (event: EventItem) => {
    const transcript = getDemoMeetingTranscript(event.subject);
    if (!transcript) return;

    setSummaryLoadingEventId(event.id);
    setSummaryErrors(current => ({ ...current, [event.id]: '' }));

    try {
      PubSub.publish('Calendar', [{
        api: getAzureOpenAIChatCompletionsUrl(),
        type: 'POST',
      }]);
      const summary = await summarizeMeetingTranscript(event.subject, transcript);
      setSummaries(current => ({ ...current, [event.id]: summary }));
    } catch (error) {
      const message = error instanceof Error
        ? error.message
        : 'Unable to generate the meeting summary.';
      setSummaryErrors(current => ({ ...current, [event.id]: message }));
    } finally {
      setSummaryLoadingEventId(null);
    }
  };

  const renderEvent = (ev: EventItem) => {
    // 先将 ISO 字符串解析为 UTC 时间，再转换成本地时间
    const startUTC = new Date(ev.start + (ev.start.endsWith('Z') ? '' : 'Z'));
    const endUTC = new Date(ev.end + (ev.end.endsWith('Z') ? '' : 'Z'));
    const demoTranscript = getDemoMeetingTranscript(ev.subject);
    const summary = summaries[ev.id];
    const summaryError = summaryErrors[ev.id];
    const isGeneratingSummary = summaryLoadingEventId === ev.id;
    return (
      <li key={ev.id} className="agenda-item">
        <div className="agenda-item-header">
          <span className="agenda-item-time">
            {startUTC.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })} - {endUTC.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
          </span>
          <div className="agenda-item-main">
            <span className="agenda-item-subject">{ev.subject}</span>
            <div className="agenda-item-subtitle">
              <span style={{display:'flex',alignItems:'center',gap:4}}>
                <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M12 2C7.03 2 3 6.03 3 11c0 5.25 7.11 10.54 8.13 11.29.53.39 1.21.39 1.74 0C13.89 21.54 21 16.25 21 11c0-4.97-4.03-9-9-9Zm0 17.88C9.09 17.07 5 13.61 5 11c0-3.87 3.13-7 7-7s7 3.13 7 7c0 2.61-4.09 6.07-7 8.88Z" fill="#666"/><circle cx="12" cy="11" r="3" fill="#666"/></svg>
                {ev.location || '—'}
              </span>
              {/* 可扩展更多子标题内容 */}
            </div>
            <div className="agenda-item-attendees">
              {renderAttendees(ev.attendees)}
            </div>
            {demoTranscript && (
              <div className="agenda-demo-summary">
                <details className="agenda-demo-transcript">
                  <summary>View synthetic transcript</summary>
                  <pre>{demoTranscript}</pre>
                </details>
                <Button
                  appearance="primary"
                  size="small"
                  icon={<Sparkle20Regular />}
                  disabled={isGeneratingSummary}
                  onClick={() => handleGenerateDemoSummary(ev)}
                >
                  {isGeneratingSummary
                    ? 'Generating summary...'
                    : summary
                      ? 'Regenerate AI summary'
                      : 'Generate AI summary'}
                </Button>
                {summaryError && (
                  <div className="agenda-summary-error" role="alert">
                    {summaryError}
                  </div>
                )}
                {summary && (
                  <div className="agenda-summary-result" aria-live="polite">
                    <strong>AI meeting summary</strong>
                    <div>{summary}</div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </li>
    );
  };

  return (
    <div className="agenda-root" id={id} key={agendaKey}>
      <h2 style={{display:'none'}}>Agenda</h2>
      {selectedUser && cardPosition && (
        <div
          style={{
            position: 'fixed',
            left: cardPosition.left,
            top: cardPosition.top,
            zIndex: 9999,
            boxShadow: '0 4px 24px rgba(0,0,0,0.18)',
            background: '#fff',
            borderRadius: 8,
          }}
          onMouseEnter={handleCardMouseEnter}
          onMouseLeave={handleCardMouseLeave}
        >
          <PersonCard
            data={selectedUser}
            onClose={handleAvatarMouseLeave}
          />
        </div>
      )}
      {groupByDay && groupedEvents
        ? Object.entries(groupedEvents).map(([day, dayEvents]) => {
            const localDate = parseLocalDateKey(day);
            return (
              <div key={day}>
                <div style={{ fontWeight: 'bold', margin: '12px 0 4px' }}>
                  {localDate.toLocaleDateString()}
                </div>
                <ul className="agenda-list">
                  {dayEvents.map(renderEvent)}
                </ul>
              </div>
            );
          })
        : (
          <ul className="agenda-list">
            {events.map(renderEvent)}
          </ul>
        )}
      {children}
    </div>
  );
};

// presenceMap 兼容更多 Graph API 返回值
const presenceMap: Record<string, PersonaPresence> = {
  available: PersonaPresence.online,
  Available: PersonaPresence.online,
  busy: PersonaPresence.busy,
  Busy: PersonaPresence.busy,
  dnd: PersonaPresence.dnd,
  DoNotDisturb: PersonaPresence.dnd,
  away: PersonaPresence.away,
  Away: PersonaPresence.away,
  brb: PersonaPresence.away,
  BeRightBack: PersonaPresence.away,
  offline: PersonaPresence.offline,
  Offline: PersonaPresence.offline,
  PresenceUnknown: PersonaPresence.none,
  OutOfOffice: PersonaPresence.away,
};

// 统一邮箱补全函数，组件作用域唯一
let defaultDomain = '';
const ensureEmail = (email: string) => {
  if (/^[^@]+@[^@]+\.[^@]+$/.test(email)) {
    if (!defaultDomain) {
      defaultDomain = email.substring(email.indexOf('@'));
    }
    return email;
  }
  return email + (defaultDomain || '@example.com');
};

export type AgendaProps = {
  groupByDay?: boolean;
  eventQuery?: string;
  id?: string;
  agendaKey?: React.Key;
  token: () => Promise<string>;
  children?: React.ReactNode;
};
