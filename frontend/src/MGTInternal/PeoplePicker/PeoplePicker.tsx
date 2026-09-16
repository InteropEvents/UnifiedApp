// 独立于MGT的PeoplePicker组件
import React, { useState, useEffect, useRef } from 'react';
import './PeoplePicker.css';

export interface Person {
  id: string;
  displayName: string;
  mail?: string;
  jobTitle?: string;
  avatarUrl?: string;
}

export interface PeoplePickerProps {
  accessToken: string;
  placeholder?: string;
  selectionMode?: 'single' | 'multiple';
  showMax?: number;
  onChange?: (selected: Person[]) => void;
  defaultSelectedUserIds?: string[];
  selectionChanged?: (e: { detail: Person[] }) => void;
}

export const PeoplePicker: React.FC<Omit<PeoplePickerProps, 'fetchPeople'> > = ({
  accessToken,
  placeholder = 'Search personnel...',
  selectionMode = 'multiple',
  showMax = 6,
  onChange,
  defaultSelectedUserIds = [],
  selectionChanged,
}) => {
  const [input, setInput] = useState('');
  const [suggestions, setSuggestions] = useState<Person[]>([]);
  const [selected, setSelected] = useState<Person[]>([]);
  const [loading, setLoading] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  // 内部人员查询逻辑，优先 /me/people，回退 /users
  async function fetchPeople(query: string, accessToken: string): Promise<Person[]> {
    const headers = {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    };
    // 1. 推荐联系人
    let people: Person[] = [];
    try {
      const res = await fetch(
        `https://graph.microsoft.com/v1.0/me/people?$search="${encodeURIComponent(query)}"`,
        { headers }
      );
      if (res.ok) {
        const data = await res.json();
        people = (data.value || []).map((p: any) => ({
          id: p.id || p.userPrincipalName || p.mail,
          displayName: p.displayName,
          mail: p.userPrincipalName,
          jobTitle: p.jobTitle,
          avatarUrl: p.userPrincipalName ? `https://graph.microsoft.com/v1.0/users/${p.userPrincipalName}/photo/$value` : undefined,
        }));
      }
    } catch {}
    // 2. 若无结果则查组织用户
    if (!people.length) {
      try {
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/users?$search="${encodeURIComponent(query)}"`,
          { headers }
        );
        if (res.ok) {
          const data = await res.json();
          people = (data.value || []).map((p: any) => ({
            id: p.id || p.userPrincipalName || p.mail,
            displayName: p.displayName,
            mail: p.userPrincipalName,
            jobTitle: p.jobTitle,
            avatarUrl: p.id ? `https://graph.microsoft.com/v1.0/users/${p.id}/photo/$value` : undefined,
          }));
        }
      } catch {}
    }
    return people;
  }

  // 异步获取头像并赋值到人员对象
  async function fetchAvatar(person: Person, accessToken: string): Promise<string> {
    if (!person.id) return '';
    try {
      const res = await fetch(`https://graph.microsoft.com/v1.0/users/${person.id}/photo/$value`, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (res.ok) {
        const blob = await res.blob();
        return URL.createObjectURL(blob);
      }
    } catch {}
    return 'https://ui-avatars.com/api/?name=' + encodeURIComponent(person.displayName);
  }

  useEffect(() => {
    if (input.trim()) {
      setLoading(true);
      fetchPeople(input, accessToken).then(async res => {
        // 批量处理头像
        const withAvatar = await Promise.all(res.slice(0, showMax).map(async p => {
          return { ...p, avatarUrl: await fetchAvatar(p, accessToken) };
        }));
        setSuggestions(withAvatar);
        setLoading(false);
      });
    } else {
      setSuggestions([]);
    }
  }, [input, accessToken, showMax]);

  const handleSelect = async (person: Person) => {
    let newSelected: Person[];
    if (selectionMode === 'single') {
      newSelected = [person];
    } else if (!selected.find(p => p.id === person.id)) {
      newSelected = [...selected, person];
    } else {
      newSelected = selected;
    }
    // 处理头像
    const withAvatar = await Promise.all(newSelected.map(async p => {
      if (!p.avatarUrl) {
        return { ...p, avatarUrl: await fetchAvatar(p, accessToken) };
      }
      return p;
    }));
    setSelected(withAvatar);
    setInput('');
    setSuggestions([]);
  };

  const handleRemove = (person: Person) => {
    const filtered = selected.filter(p => p.id !== person.id);
    setSelected(filtered);
  };

  // 初始化默认选中人员，只执行一次
  const didInitRef = useRef(false);
  // 记录上一次的 defaultSelectedUserIds
  const prevDefaultIdsRef = useRef<string[]>([]);
  // 标记是否是初始化选中
  const isInitRef = useRef(true);

  useEffect(() => {
    if (isInitRef.current) {
      // 初始化时不触发回调
      isInitRef.current = false;
      return;
    }
    onChange?.(selected);
    selectionChanged?.({ detail: selected });
  }, [selected, onChange, selectionChanged]);

  useEffect(() => {
    // 只有 defaultSelectedUserIds 发生变化时才初始化
    const idsStr = (defaultSelectedUserIds || []).join(',');
    const prevIdsStr = prevDefaultIdsRef.current.join(',');
    if (idsStr !== prevIdsStr && defaultSelectedUserIds && defaultSelectedUserIds.length > 0) {
      prevDefaultIdsRef.current = [...defaultSelectedUserIds];
      (async () => {
        const users = await Promise.all(defaultSelectedUserIds.map(async id => {
          const headers = {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          };
          let user: Person | null = null;
          try {
            const res = await fetch(`https://graph.microsoft.com/v1.0/users/${id}`, { headers });
            if (res.ok) {
              const data = await res.json();
              user = {
                id: data.id || data.userPrincipalName || data.mail,
                displayName: data.displayName,
                mail: data.mail,
                jobTitle: data.jobTitle,
                avatarUrl: await fetchAvatar(data, accessToken),
              };
            }
          } catch {}
          return user;
        }));
        setSelected(users.filter(Boolean) as Person[]);
      })();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [defaultSelectedUserIds, accessToken]);

  return (
    <div className="people-picker">
      <div className="selected-list">
        {selected.map(person => (
          <span className="selected-person" key={person.id}>
            <img src={person.avatarUrl || ''} alt="avatar" className="avatar" />
            {person.displayName}
            <button className="remove-btn" onClick={() => handleRemove(person)}>&times;</button>
          </span>
        ))}
      </div>
      <input
        ref={inputRef}
        type="text"
        className="people-picker-input"
        placeholder={placeholder}
        value={input}
        onChange={e => setInput(e.target.value)}
        disabled={selectionMode === 'single' && selected.length === 1}
      />
      {loading && <div className="loading">加载中...</div>}
      {input && suggestions.length > 0 && (
        <ul className="suggestions-list">
          {suggestions.map(person => (
            <li key={person.id} onClick={() => handleSelect(person)}>
              <img src={person.avatarUrl || ''} alt="avatar" className="avatar" />
              <span>{person.displayName}</span>
              <span className="mail">{person.mail}</span>
              <span className="jobTitle">{person.jobTitle}</span>
            </li>
          ))}
        </ul>
      )}
      {input && !loading && suggestions.length === 0 && (
        <div className="no-results">无结果</div>
      )}
    </div>
  );
};
