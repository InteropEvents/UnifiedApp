import React, { useRef, useEffect, useState } from 'react';
import styles from './PersonCard.module.css';
import { usePersonData, PersonData, PersonFetcher } from '../personcard/usePersonData';
import { MailIcon, PhoneIcon, MobileIcon, LocationIcon } from './ContactIcons';
import PubSub from 'pubsub-js';

export interface PersonCardProps {
  userId?: string;
  data?: PersonData;
  fetcher?: PersonFetcher;
  style?: React.CSSProperties;
  className?: string;
  onClose?: () => void;
}

const TABS = [
  { label: 'Profile', key: 'profile' },
  { label: 'Emails', key: 'mail' },
  { label: 'Files', key: 'files' },
  { label: 'About', key: 'about' }
];

export const PersonCard: React.FC<PersonCardProps> = ({
  userId,
  data,
  fetcher,
  style,
  className,
  onClose
}) => {
  const { person, loading, error } = usePersonData({ userId, data, fetcher });
  const [tab, setTab] = useState(0);
  const [showMoreMail, setShowMoreMail] = useState(false);
  const [showMoreFile, setShowMoreFile] = useState(false);
  const overlayRef = useRef<HTMLDivElement>(null);
  const cardRef = useRef<HTMLDivElement>(null);
  const [mails, setMails] = useState<any[]>([]);
  const [loadingMails, setLoadingMails] = useState(false);
  const [mailError, setMailError] = useState<string | null>(null);
  const [files, setFiles] = useState<any[]>([]);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [fileError, setFileError] = useState<string | null>(null);
  const [profile, setProfile] = useState<any>(null);
  const [loadingProfile, setLoadingProfile] = useState(false);
  const [profileError, setProfileError] = useState<string | null>(null);

  // ESC关闭、Tab循环
  useEffect(() => {
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose?.();
      else if (e.key === 'Tab') {
        const focusable = cardRef.current?.querySelectorAll<HTMLElement>(
          'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
        );
        if (!focusable || focusable.length === 0) return;
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (e.shiftKey) {
          if (document.activeElement === first) {
            last.focus();
            e.preventDefault();
          }
        } else {
          if (document.activeElement === last) {
            first.focus();
            e.preventDefault();
          }
        }
      }
    };
    document.addEventListener('keydown', handleKey);
    return () => document.removeEventListener('keydown', handleKey);
  }, [onClose]);

  // 点击遮罩关闭
  useEffect(() => {
    const handleClick = (e: MouseEvent) => {
      if (overlayRef.current && e.target === overlayRef.current) {
        onClose?.();
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [onClose]);

  // 打印person参数到控制台
  useEffect(() => {
  }, [person]);

  // 查询组织信息
  useEffect(() => {
    if (!person?.id) return;
    let canceled = false;
    async function fetchOrganizations() {
      if (!person?.id) return;
      try {
        const { getGraphSdkClient } = await import('../../services/graphSdkClient');
        const client = await getGraphSdkClient();
        const apiUrl = `https://graph.microsoft.com/v1.0/users/${person.id}/memberOf`;
        PubSub.publish('Calendar', [{ api: apiUrl, type: 'GET' }]);
        const res = await client.api(`/users/${person.id}/memberOf`).get();
        const orgs = (res.value || [])
          .filter((item: any) => item['@odata.type'] === '#microsoft.graph.group' || item['@odata.type'] === '#microsoft.graph.directoryRole')
          .map((item: any) => item.displayName || item.roleTemplateId || item.id);
        if (!canceled && person) {
          person.organizations = orgs;
        }
      } catch (e) {
        // 静默忽略，不打印警告
      }
    }
    fetchOrganizations();
    return () => { canceled = true; };
  }, [person?.id]);

  // 获取与该用户相关的邮件
  useEffect(() => {
    async function fetchMails() {
      setLoadingMails(true);
      setMailError(null);
      setMails([]);
      if (!person?.mail) return;
      try {
        const { getGraphSdkClient } = await import('../../services/graphSdkClient');
        const client = await getGraphSdkClient();
        const apiUrl = `https://graph.microsoft.com/v1.0/me/messages?$search=\"from:${person.mail}\"`;
        PubSub.publish('Calendar', [{ api: apiUrl, type: 'GET' }]);
        const res = await client.api(`/me/messages?$search=\"from:${person.mail}\"`).get();
        setMails(res.value || []);
      } catch (e: any) {
        setMailError(e.message || '获取邮件失败');
      } finally {
        setLoadingMails(false);
      }
    }
    fetchMails();
  }, [person?.mail]);

  // 获取与该用户相关的文件（最近共享、最近使用）
  useEffect(() => {
    async function fetchFiles() {
      setLoadingFiles(true);
      setFileError(null);
      setFiles([]);
      if (!person?.mail) return;
      try {
        const { getGraphSdkClient } = await import('../../services/graphSdkClient');
        const client = await getGraphSdkClient();
        const sharedApi = `https://graph.microsoft.com/v1.0/me/insights/shared?$filter=lastshared/sharedby/address eq '${person.mail}'`;
        const usedApi = 'https://graph.microsoft.com/v1.0/me/insights/used';
        PubSub.publish('Calendar', [{ api: sharedApi, type: 'GET' }]);
        const sharedRes = await client.api(`/me/insights/shared?$filter=lastshared/sharedby/address eq '${person.mail}'`).get();
        PubSub.publish('Calendar', [{ api: usedApi, type: 'GET' }]);
        const usedRes = await client.api('/me/insights/used').get();
        // 合并并去重，优先展示最近共享
        const allFiles = [...(sharedRes.value || []), ...(usedRes.value || [])];
        const fileMap = new Map();
        allFiles.forEach((item: any) => {
          const file = item.resourceReference || item.resourceVisualization || item.resource || item;
          if (file && file.id && !fileMap.has(file.id)) {
            fileMap.set(file.id, {
              id: file.id,
              name: file.displayName || file.title || file.name,
              url: file.webUrl || file.url,
              lastModifiedDateTime: file.lastModifiedDateTime || item.lastUsed?.lastAccessedDateTime || item.lastShared?.lastSharedDateTime,
              size: file.size
            });
          }
        });
        setFiles(Array.from(fileMap.values()).slice(0, 5));
      } catch (e: any) {
        setFileError(e.message || '获取文件失败');
      } finally {
        setLoadingFiles(false);
      }
    }
    fetchFiles();
  }, [person?.mail]);

  // 获取用户profile信息（skills、languages、positions、educationalActivities、interests等）
  useEffect(() => {
    async function fetchProfile() {
      setLoadingProfile(true);
      setProfileError(null);
      setProfile(null);
      if (!person?.id) return;
      try {
        const { getGraphSdkClient } = await import('../../services/graphSdkClient');
        const client = await getGraphSdkClient();
        const apiUrl = `https://graph.microsoft.com/beta/users/${person.id}/profile`;
        PubSub.publish('Calendar', [{ api: apiUrl, type: 'GET' }]);
        const res = await client.api(`/users/${person.id}/profile`).version('beta').get();
        setProfile(res);
      } catch (e: any) {
        setProfileError(e.message || '获取profile失败');
      } finally {
        setLoadingProfile(false);
      }
    }
    fetchProfile();
  }, [person?.id]);

  // tab内容
  const tabPanels = [
    // Profile
    <>
      <div className={styles.section}>
        <div className={styles.sectionTitle}>Contact</div>
        <div className={styles.sectionContent}>
          {person?.mail && (
            <div className={styles.contactItem} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <MailIcon />
              <a href={`mailto:${person.mail}`} className={styles.link}>{person.mail}</a>
            </div>
          )}
          {person?.phones?.[0] && (
            <div className={styles.contactItem} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <PhoneIcon />
              <a href={`tel:${person.phones[0]}`} className={styles.link}>{person.phones[0]}</a>
            </div>
          )}
          {person?.phones?.[1] && (
            <div className={styles.contactItem} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <MobileIcon />
              <a href={`tel:${person.phones[1]}`} className={styles.link}>{person.phones[1]}</a>
            </div>
          )}
          {person?.address && (
            <div className={styles.contactItem} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <LocationIcon />
              <span>{person.address}</span>
            </div>
          )}
        </div>
      </div>
      <div className={styles.section}>
        <div className={styles.sectionTitle}>Presence</div>
        <div className={styles.sectionContent}>{person?.presenceText || '—'}</div>
      </div>
    </>,
    // Mail
    <div className={styles.section}>
      <div className={styles.sectionTitle}>Emails</div>
      <div className={styles.sectionContent} style={{ maxHeight: 220, overflowY: mails.length > 5 ? 'auto' : 'visible' }}>
        {loadingMails ? (
          <div>Loading mails…</div>
        ) : mailError ? (
          <div style={{ color: 'red' }}>{mailError}</div>
        ) : mails.length > 0 ? (
          <div>
            {mails.slice(0, 5).map((mail, i) => (
              <div key={mail.id || i} className={styles.mailItem} style={{ marginBottom: 8, paddingLeft: 0 }}>
                <div><strong>{mail.subject || '(No subject)'}</strong></div>
                <div style={{ fontSize: 12, color: '#888' }}>
                  {mail.from?.emailAddress?.name || mail.from?.emailAddress?.address || 'Unknown sender'}
                  {' | '}
                  {mail.receivedDateTime ? new Date(mail.receivedDateTime).toLocaleString() : ''}
                </div>
              </div>
            ))}
          </div>
        ) : (
          <div>No related mails</div>
        )}
      </div>
    </div>,
    // Files
    <div className={styles.section}>
      <div className={styles.sectionTitle}>Files</div>
      <div className={styles.sectionContent} style={{ maxHeight: 220, overflowY: files.length > 5 ? 'auto' : 'visible' }}>
        {loadingFiles ? (
          <div>Loading files…</div>
        ) : fileError ? (
          <div style={{ color: 'red' }}>{fileError}</div>
        ) : files.length > 0 ? (
          <ul style={{ paddingLeft: 0, margin: 0, listStyle: 'none' }}>
            {files.map((f, i) => {
              // 用正则从url中提取文件名
              let fileName = f.name || f.displayName || f.title;
              if (!fileName && f.url) {
                const match = decodeURIComponent(f.url).match(/([^/\\?#]+)(?=[?#]|$)/);
                fileName = match ? match[1] : f.url;
              }
              return (
                <li key={f.id || i} className={styles.fileItem} style={{ marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span style={{ fontSize: 18 }}>📄</span>
                  {f.url ? (
                    <a href={f.url} target="_blank" rel="noopener noreferrer" className={styles.link}>{fileName || '未知文件'}</a>
                  ) : (
                    <span>{fileName || '未知文件'}</span>
                  )}
                  {f.lastModifiedDateTime && (
                    <span style={{ fontSize: 12, color: '#888' }}>{new Date(f.lastModifiedDateTime).toLocaleString()}</span>
                  )}
                  {typeof f.size === 'number' && (
                    <span style={{ fontSize: 12, color: '#888' }}>{(f.size / 1024).toFixed(1)} KB</span>
                  )}
                </li>
              );
            })}
          </ul>
        ) : 'No files'}
      </div>
    </div>,
    // About
    <div className={styles.section}>
      <div className={styles.sectionTitle}>About</div>
      <div className={styles.sectionContent}>
        {loadingProfile ? (
          <div>Loading profile…</div>
        ) : profileError ? (
          <div style={{ color: 'red' }}>{profileError}</div>
        ) : profile ? (
          <div>
            {/* 技能 */}
            {profile.skills && profile.skills.length > 0 && (
              <div style={{ marginBottom: 8 }}>
                <strong>Skills:</strong> {profile.skills.map((s: any) => s.displayName).join(', ')}
              </div>
            )}
            {/* 语言 */}
            {profile.languages && profile.languages.length > 0 && (
              <div style={{ marginBottom: 8 }}>
                <strong>Languages:</strong> {profile.languages.map((l: any) => l.displayName).join(', ')}
              </div>
            )}
            {/* 工作经历 */}
            {profile.positions && profile.positions.length > 0 && (
              <div style={{ marginBottom: 8 }}>
                <strong>Work Experience:</strong>
                <ul style={{ margin: 0, paddingLeft: 0, listStyle: 'none' }}>
                  {profile.positions.map((p: any, i: number) => (
                    <li key={i}>{
                      (p.company?.displayName || p.company?.name || p.company || '') +
                      (p.jobTitle ? ` - ${typeof p.jobTitle === 'object' ? (p.jobTitle.displayName || p.jobTitle.name || JSON.stringify(p.jobTitle)) : p.jobTitle}` : '')
                    }</li>
                  ))}
                </ul>
              </div>
            )}
            {/* 教育经历 */}
            {profile.educationalActivities && profile.educationalActivities.length > 0 && (
              <div style={{ marginBottom: 8 }}>
                <strong>Education:</strong>
                <ul style={{ margin: 0, paddingLeft: 0, listStyle: 'none' }}>
                  {profile.educationalActivities.map((e: any, i: number) => {
                    let institution = e.institution?.displayName || e.institution?.name || e.institution || '';
                    let program = '';
                    if (e.program) {
                      if (typeof e.program === 'object') {
                        program = e.program.displayName || e.program.name || '';
                      } else {
                        program = e.program;
                      }
                    }
                    // 如果 program 为空或全是空字符串，则不显示
                    if (!institution && !program) return null;
                    return (
                      <li key={i}>{institution}{program ? ` - ${program}` : ''}</li>
                    );
                  })}
                </ul>
              </div>
            )}
            {/* 兴趣 */}
            {profile.interests && profile.interests.length > 0 && (
              <div style={{ marginBottom: 8 }}>
                <strong>Interests:</strong> {profile.interests.map((it: any) => it.displayName).join(', ')}
              </div>
            )}
            {/* 个人简介 */}
            {profile.aboutMe && (
              <div style={{ marginBottom: 8 }}>
                <strong>About Me:</strong> {profile.aboutMe}
              </div>
            )}
            {/* 没有任何信息时 */}
            {!(profile.skills?.length || profile.languages?.length || profile.positions?.length || profile.educationalActivities?.length || profile.interests?.length || profile.aboutMe) && (
              <div>No about info</div>
            )}
          </div>
        ) : (
          person?.about || person?.experience || 'No about info'
        )}
      </div>
    </div>
  ];

  return (
    <div className={styles.card} ref={cardRef} style={style} tabIndex={0} aria-label="Person Card">
      <button className={styles.close} onClick={onClose} aria-label="Close">×</button>
      <div className={styles.header}>
        <div className={styles.avatarWrap}>
          {person?.photo ? (
            <img src={person.photo} alt={person.name || 'avatar'} className={styles.avatar} />
          ) : (
            <div className={styles.avatarPlaceholder}>?</div>
          )}
        </div>
        <div className={styles.main}>
          <div className={styles.name}>{person?.name || '—'}</div>
          <div className={styles.title}>{person?.jobTitle || ''}</div>
          <div className={styles.department}>{person?.department || ''}</div>
          <div className={styles.mail}>{person?.mail || ''}</div>
          <div className={styles.presence}>{person?.presenceText || ''}</div>
        </div>
      </div>
      <div className={styles.tabs}>
        {TABS.map((t, i) => (
          <button
            key={t.key}
            className={styles.tab + (tab === i ? ' ' + styles.tabActive : '')}
            onClick={() => setTab(i)}
            tabIndex={0}
            aria-selected={tab === i}
          >
            {t.label}
          </button>
        ))}
      </div>
      <div className={styles.tabPanel}>{tabPanels[tab]}</div>
      {loading && <div className={styles.loading}>Loading…</div>}
      {error && <div className={styles.error}>Failed: {error.message}</div>}
      {!person && !loading && !error && <div className={styles.empty}>No data</div>}
    </div>
  );
};
