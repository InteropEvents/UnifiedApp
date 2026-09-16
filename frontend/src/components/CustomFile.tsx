import * as React from 'react';
import { Button, Tooltip } from '@fluentui/react-components';
import { Eye24Regular } from '@fluentui/react-icons';
import { Icon } from '@fluentui/react/lib/Icon';
import { FileIconType, getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import { getFileExtension, getFileIconKind } from './fileIcon';
import '../MGTInternal/FileList/FileList.css';

const iconStyle = { width: 24, height: 24, flexShrink: 0 };

function getFileIcon(file: any): JSX.Element {
  const kind = getFileIconKind(file);
  const extension = getFileExtension(file);
  const iconProps = kind === 'folder'
    ? getFileTypeIconProps({ type: FileIconType.folder, size: 24 })
    : extension
      ? getFileTypeIconProps({ extension, size: 24 })
      : getFileTypeIconProps({ type: FileIconType.genericFile, size: 24 });

  return <Icon {...iconProps} aria-hidden style={iconStyle} />;
}

export const CustomFile = ({ fileDetails, onPreview }: { fileDetails: any, onPreview: () => void }) => (
  <div className="mgt-custom-file-row">
    <a
      className="mgt-custom-file-link"
      href={fileDetails.webUrl}
      target="_blank"
      rel="noopener noreferrer"
      aria-label={`Open ${fileDetails.name}`}
    >
      <span className="mgt-custom-file-icon">{getFileIcon(fileDetails)}</span>
      <span className="mgt-custom-file-name" title={fileDetails.name}>{fileDetails.name}</span>
    </a>
    <Tooltip content={`Preview ${fileDetails.name}`} relationship="description">
      <Button
        className="mgt-custom-file-preview"
        appearance="subtle"
        size="small"
        icon={<Eye24Regular />}
        onClick={onPreview}
        aria-label={`Preview ${fileDetails.name}`}
      />
    </Tooltip>
  </div>
);
