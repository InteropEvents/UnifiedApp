// FileList.tsx
import React, { useEffect, useState, useRef } from 'react';
import { Button } from '@fluentui/react-components';
import { ArrowUpload20Regular } from '@fluentui/react-icons';
import './FileList.css';
import { fetchFiles, GraphFile, getGraphClient } from './graphService';
import { FileListItem } from './FileListItem';
import PubSub from 'pubsub-js';

export interface FileListProps {
  accessToken: string;
  driveId?: string;
  siteId?: string;
  groupId?: string; // 支持 groupId
  itemPath?: string;
  pageSize?: number;
  className?: string;
  enableFileUpload?: boolean;
  disableOpenOnClick?: boolean;
  onFileClick?: (file: GraphFile) => void;
  children?: React.ReactNode;
  onLoadingChange?: (loading: boolean) => void; // 新增回调
}

export const FileList = React.forwardRef<HTMLDivElement, FileListProps>((props, ref) => {
  const {
    accessToken,
    driveId,
    siteId,
    groupId,
    itemPath,
    pageSize,
    className,
    enableFileUpload,
    disableOpenOnClick,
    onFileClick,
    onLoadingChange
  } = props;

  const [files, setFiles] = useState<GraphFile[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const loadFiles = async () => {
    setLoading(true);
    setError(null);
    try {
      // 优先使用 groupId 作为 driveId
      const endpoint = (() => {
        if (groupId) return `/groups/${groupId}/drive/root${itemPath ? ':/' + itemPath + ':/children' : '/children'}`;
        if (driveId) return `/drives/${driveId}/root${itemPath ? ':/' + itemPath + ':/children' : '/children'}`;
        if (siteId) return `/sites/${siteId}/drive/root${itemPath ? ':/' + itemPath + ':/children' : '/children'}`;
        return `/me/drive/root${itemPath ? ':/' + itemPath + ':/children' : '/children'}`;
      })();
      const apiUrl = `https://graph.microsoft.com/v1.0${endpoint}`;
      PubSub.publish('Calendar', [{ api: apiUrl, type: 'GET' }]);
      const result = await fetchFiles({
        accessToken,
        driveId: groupId || driveId,
        siteId,
        itemPath,
        pageSize
      });
      setFiles(result);
    } catch (err: any) {
      setError(err.message || 'Failed to load files');
    }
    setLoading(false);
  };

  useEffect(() => {
    loadFiles();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accessToken, driveId, siteId, groupId, itemPath, pageSize]);

  useEffect(() => {
    if (onLoadingChange) {
      onLoadingChange(loading);
    }
  }, [loading, onLoadingChange]);

  const getUploadEndpoint = ({ groupId, driveId, siteId, itemPath, file }: {
    groupId?: string;
    driveId?: string;
    siteId?: string;
    itemPath?: string;
    file: File;
  }) => {
    const safeItemPath = itemPath ? encodeURIComponent(itemPath) : '';
    const safeFileName = encodeURIComponent(file.name);
    if (groupId) {
      return `/groups/${groupId}/drive/root:/${safeItemPath}${safeItemPath ? '/' : ''}${safeFileName}:/content`;
    } else if (driveId) {
      return `/drives/${driveId}/root:/${safeItemPath}${safeItemPath ? '/' : ''}${safeFileName}:/content`;
    } else if (siteId) {
      return `/sites/${siteId}/drive/root:/${safeItemPath}${safeItemPath ? '/' : ''}${safeFileName}:/content`;
    } else {
      return `/me/drive/root:/${safeItemPath}${safeItemPath ? '/' : ''}${safeFileName}:/content`;
    }
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const filesToUpload = event.target.files;
    if (!filesToUpload || filesToUpload.length === 0) return;
    for (let i = 0; i < filesToUpload.length; i++) {
      const file = filesToUpload[i];
      try {
        const endpoint = getUploadEndpoint({ groupId, driveId, siteId, itemPath, file });
        const apiUrl = `https://graph.microsoft.com/v1.0${endpoint}`;
        PubSub.publish('Calendar', [{ api: apiUrl, type: 'PUT' }]);
        const client = getGraphClient(accessToken);
        await client.api(endpoint).put(file);
      } catch (err) {
        setError(`Upload failed: ${file.name}`);
      }
    }
    loadFiles();
  };

  return (
    <div ref={ref} className={`mgt-file-list ${className || ''}`}>
      {enableFileUpload && (
        <div className="mgt-file-list__upload">
          <input ref={fileInputRef} type="file" multiple hidden onChange={handleFileUpload} />
          <Button
            appearance="primary"
            icon={<ArrowUpload20Regular />}
            onClick={() => fileInputRef.current?.click()}
          >
            Upload file
          </Button>
        </div>
      )}
      <div className="mgt-file-list__content">
        {props.children}
        {loading && <div className="mgt-file-list__loading">Loading...</div>}
        {error && <div className="mgt-file-list__error">{error}</div>}
        {!loading && !error && files.length === 0 && (
          <div className="mgt-file-list__empty">No files</div>
        )}
        {!loading && !error && files.map(file => (
          props.children
            ? React.Children.map(props.children, child =>
                React.isValidElement(child)
                  ? React.cloneElement(child as any, { dataContext: { file } })
                  : child
              )
            : (
              <FileListItem
                key={file.id}
                file={file}
                onClick={onFileClick}
                disableOpenOnClick={disableOpenOnClick}
              />
            )
        ))}
      </div>
    </div>
  );
});
