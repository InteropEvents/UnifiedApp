import { useEffect, useRef, useState } from 'react';

export interface PersonData {
  id?: string;
  name?: string;
  photo?: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  phones?: string[];
  chat?: string;
  profileUrl?: string;
  presenceText?: string;
  recentMails?: { id?: string; subject: string; url: string }[];
  files?: { id?: string; name: string; url: string }[];
  skills?: string[];
  experience?: string;
  address?: string;
  mobile?: string;
  organizations?: string[];
  userPrincipalName?: string;
  proxyAddresses?: string[];
  about?: string; // <-- Add about field here
  languages?: string[];
}

export type PersonFetcher = (userId: string) => Promise<PersonData>;

const cache = new Map<string, PersonData>();

export function usePersonData({ userId, data, fetcher }: {
  userId?: string;
  data?: PersonData;
  fetcher?: PersonFetcher;
}) {
  const [person, setPerson] = useState<PersonData | undefined>(data);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<Error | null>(null);
  const lastId = useRef<string | undefined>();

  useEffect(() => {
    if (data) {
      setPerson(data);
      setLoading(false);
      setError(null);
      return;
    }
    if (!userId || !fetcher) {
      setPerson(undefined);
      setLoading(false);
      setError(null);
      return;
    }
    if (cache.has(userId)) {
      setPerson(cache.get(userId));
      setLoading(false);
      setError(null);
      return;
    }
    setLoading(true);
    setError(null);
    lastId.current = userId;
    fetcher(userId)
      .then(res => {
        if (lastId.current !== userId) return;
        cache.set(userId, res);
        setPerson(res);
        setLoading(false);
      })
      .catch(e => {
        if (lastId.current !== userId) return;
        setError(e instanceof Error ? e : new Error(String(e)));
        setLoading(false);
      });
  }, [userId, data, fetcher]);

  return { person, loading, error };
}
