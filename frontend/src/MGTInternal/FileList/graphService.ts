// graphService.ts
import { Client } from '@microsoft/microsoft-graph-client';

export interface GraphFile {
  id: string;
  name: string;
  webUrl: string;
  lastModifiedDateTime: string;
  size: number;
  file?: { mimeType: string };
  folder?: {};
  [key: string]: any;
}

export interface FileListOptions {
  accessToken: string;
  driveId?: string;
  siteId?: string;
  groupId?: string;
  itemPath?: string;
  pageSize?: number;
}

export function getGraphClient(accessToken: string): Client {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
}

export async function fetchFiles(options: FileListOptions): Promise<GraphFile[]> {
  const { accessToken, driveId, siteId, groupId, itemPath, pageSize = 20 } = options;
  const client = getGraphClient(accessToken);
  let endpoint = '';
  if (driveId) {
    endpoint = `/groups/${driveId}/drive/root${itemPath ? `:/${itemPath}:/children` : '/children'}`;
  } else if (siteId) {
    endpoint = `/sites/${siteId}/drive/root${itemPath ? `:/${itemPath}:/children` : '/children'}`;
  } else if (driveId) {
    endpoint = `/drives/${driveId}/root${itemPath ? `:/${itemPath}:/children` : '/children'}`;
  } else {
    endpoint = `/me/drive/root${itemPath ? `:/${itemPath}:/children` : '/children'}`;
  }

  const response = await client.api(endpoint).top(pageSize).get();
  return response.value as GraphFile[];
}

export async function uploadFile(options: {
  accessToken: string;
  driveId?: string;
  groupId?: string;
  itemPath?: string;
  file: File;
}) {
  const { accessToken, driveId, groupId, itemPath, file } = options;
  const client = getGraphClient(accessToken);
  const usedDriveId = groupId || driveId;
  if (!usedDriveId || !itemPath || !file) throw new Error('缺少必要参数');
  const safeItemPath = itemPath ? encodeURIComponent(itemPath) : '';
  const safeFileName = encodeURIComponent(file.name);
  const endpoint = `/groups/${usedDriveId}/root:/${safeItemPath}${safeItemPath ? '/' : ''}${safeFileName}:/content`;
  await client.api(endpoint).put(file);
}
