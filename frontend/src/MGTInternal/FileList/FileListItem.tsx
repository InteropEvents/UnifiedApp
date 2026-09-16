// FileListItem.tsx
import React from 'react';
import './FileList.css';
import { GraphFile } from './graphService';

function formatSize(size: number): string {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  if (size < 1024 * 1024 * 1024) return `${(size / 1024 / 1024).toFixed(1)} MB`;
  return `${(size / 1024 / 1024 / 1024).toFixed(1)} GB`;
}

function formatDate(date: string): string {
  return new Date(date).toLocaleString();
}

function getFileIcon(file: GraphFile): string {
  if (file.folder) return '📁';
  if (file.file?.mimeType?.startsWith('image/')) return '🖼️';
  if (file.file?.mimeType?.startsWith('video/')) return '🎬';
  if (file.file?.mimeType?.startsWith('audio/')) return '🎵';
  if (file.file?.mimeType?.includes('pdf')) return '📄';
  return '📄';
}

export interface FileListItemProps {
  file: GraphFile;
  onClick?: (file: GraphFile) => void;
  disableOpenOnClick?: boolean;
}

export const FileListItem: React.FC<FileListItemProps> = ({ file, onClick, disableOpenOnClick }) => {
  const handleClick = () => {
    if (disableOpenOnClick) {
      onClick?.(file);
    } else {
      window.open(file.webUrl, '_blank');
    }
  };

  return (
    <div className="mgt-file-list-item" onClick={handleClick} tabIndex={0} role="button">
      <span className="mgt-file-list-item__icon">{getFileIcon(file)}</span>
      <span className="mgt-file-list-item__name">{file.name}</span>
      <span className="mgt-file-list-item__date">{formatDate(file.lastModifiedDateTime)}</span>
      <span className="mgt-file-list-item__size">{file.folder ? '' : formatSize(file.size)}</span>
    </div>
  );
};
