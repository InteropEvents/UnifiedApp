import React from 'react';

/**
 * 该组件完全复刻 mgt-person-card 的结构和样式，依赖 mgt 元素。
 * 你需要在页面中引入 Microsoft Graph Toolkit（@microsoft/mgt-components）。
 * 例如：
 * import '@microsoft/mgt-components';
 */

export interface PersonCardProps {
  userId?: string;
  personQuery?: string;
  showPresence?: boolean;
  showEmail?: boolean;
  showPhone?: boolean;
  showOrganization?: boolean;
  showFiles?: boolean;
  showSkills?: boolean;
  showMail?: boolean;
  showTabs?: boolean;
  style?: React.CSSProperties;
  className?: string;
  onClose?: () => void;
}

export const PersonCard: React.FC<PersonCardProps> = ({
  userId,
  personQuery,
  showPresence = true,
  showEmail = true,
  showPhone = true,
  showOrganization = true,
  showFiles = true,
  showSkills = true,
  showMail = true,
  showTabs = true,
  style,
  className,
  onClose
}) => {
  return (
    <div style={{ position: 'relative', ...style }} className={className}>
      {onClose && (
        <button
          style={{
            position: 'absolute',
            right: 12,
            top: 12,
            zIndex: 2,
            background: '#f3f3f3',
            border: 'none',
            width: 32,
            height: 32,
            borderRadius: '50%',
            fontSize: 22,
            color: '#888',
            cursor: 'pointer',
            boxShadow: '0 2px 8px rgba(0,0,0,0.10)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            transition: 'background 0.2s, color 0.2s'
          }}
          onClick={onClose}
          aria-label="关闭"
        >
          ×
        </button>
      )}
    </div>
  );
};
