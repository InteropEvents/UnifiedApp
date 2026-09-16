// src/MGTInternal/Providers/MessageProps.ts
// 独立于 mgt 的消息组件 props 类型定义

export interface MessageSender {
  emailAddress?: {
    name?: string | null;
    address?: string | null;
  } | null;
}

export interface EmailMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  receivedDateTime: string;
  isRead: boolean;
  sender?: MessageSender | null;
}

export interface MessageProps {
  email: EmailMessage;
}
