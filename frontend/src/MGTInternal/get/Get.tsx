import React, { useEffect, useRef, useState } from 'react';
import { Client } from '@microsoft/microsoft-graph-client';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig } from '../login/msalConfig';
import { getToken } from '../components/getToken';

const msalInstance = new PublicClientApplication(msalConfig);

// 实际MSAL获取token实现
async function getAccessToken(scopes?: string[]): Promise<string> {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts || accounts.length === 0) throw new Error('No account found');
  return await getToken(msalInstance, accounts[0], scopes || ['User.Read']);
}

interface GetProps {
  instanceKey?: React.Key;
  resource: string;
  scopes?: string[];
  pollingRate?: number;
  onDataChange?: (data: any) => void;
  children?: (data: any, loading: boolean, error: any) => React.ReactNode;
}

export const CustomGet: React.FC<GetProps> = ({
  instanceKey,
  resource,
  scopes,
  pollingRate = 0,
  onDataChange,
  children,
}) => {
  const [data, setData] = useState<any>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<any>(null);
  const intervalRef = useRef<NodeJS.Timeout | null>(null);
  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      const token = await getAccessToken(scopes);
      const options = {
        authProvider: (done: any) => {
          done(null, token);
        },
      };
      const client = Client.init(options);
      const response = await client.api(resource).get();
      setData(response);
      if (onDataChange) onDataChange(response);
    } catch (err) {
      setError(err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    setData(null);
    setError(null);
    setLoading(true);
    fetchData();
    if (pollingRate && pollingRate > 0) {
      intervalRef.current = setInterval(fetchData, pollingRate);
      return () => {
        if (intervalRef.current) clearInterval(intervalRef.current);
      };
    }
    return () => {};
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [resource, pollingRate, instanceKey]);

  if (children) {
    return <>{children(data, loading, error)}</>;
  }
  return null;
};
