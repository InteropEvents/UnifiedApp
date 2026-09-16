import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { Messages } from '../components/Messages';
import { MailContent } from '../components/MailContent';
import { Loading } from '../components/Loading';
import { Client } from '@microsoft/microsoft-graph-client';
import {
  SelectTabData,
  SelectTabEvent,
  Tab,
  TabList,
  TabValue,
  Button,
} from '@fluentui/react-components';
import { Mail16Regular, MailRead16Regular } from '@fluentui/react-icons';
import { useOutlookPageStyles } from '../styles/Styles';
import { CustomGet } from '../MGTInternal/get/Get';
import { getToken } from '../MGTInternal/components/getToken';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig } from '../MGTInternal/login/msalConfig';

const msalInstance = new PublicClientApplication(msalConfig);
const account = msalInstance.getAllAccounts()[0] || null;

interface MailFolder {
  id: string;
  displayName: string;
  unreadItemCount?: number;
  totalItemCount?: number;
}

interface EmailMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  receivedDateTime: string;
  isRead: boolean;
  sender: {
    emailAddress: {
      name: string;
      address: string;
    }
  };
}

const FolderProvider = React.memo(({
  onFoldersLoaded
}: {
  onFoldersLoaded: (folders: MailFolder[]) => void
}) => {
  const handleFoldersLoaded = (folderData: any): MailFolder | undefined => {
    var name = folderData.displayName;
    if (['Conversation History', 'Clutter', 'Sync Issues'].includes(name)) {
      return;
    }
    return folderData;
  };
  const handleDataChange = (data: any) => {
    if (!data || !data.value) {
      return;
    }
    const newFolders = data.value
      .map((folderData: any) => handleFoldersLoaded(folderData))
      .filter((folder: any): folder is MailFolder => folder !== undefined);
    onFoldersLoaded(newFolders);
  };

  return (
    <CustomGet
      instanceKey={'folders'}
      resource="/me/mailFolders"
      scopes={['Mail.Read']}
      pollingRate={60000} // Poll every minute
      onDataChange={handleDataChange}
    />
  );
});

const MessageProvider = React.memo(({
  folderId,
  onMessagesLoaded
}: {
  folderId: string;
  onMessagesLoaded: (messages: EmailMessage[]) => void;
}) => {
  const handleDataChange = (data: any) => {
    if (!data || !data.value) {
      return;
    }
    onMessagesLoaded(data.value);
  };

  return (
    <CustomGet
      instanceKey={`messages-${folderId}`}
      resource={`/me/mailFolders/${folderId}/messages?$orderby=receivedDateTime DESC&$top=20`}
      scopes={['Mail.Read']}
      pollingRate={60 * 1000}
      onDataChange={handleDataChange}
    >
      {(data, loading, error) =>
        loading ? (
          <Loading template="loading" message={`Loading messages...`} />
        ) : null
      }
    </CustomGet>
  );
});

export const OutlookPage: React.FunctionComponent = () => {
  const styles = useOutlookPageStyles();
  const [mailFolders, setMailFolders] = React.useState<MailFolder[]>([]);
  const [selectedTab, setSelectedTab] = React.useState<TabValue>('');
  const [selectedEmail, setSelectedEmail] = React.useState<any>(null);
  const [messages, setMessages] = React.useState<EmailMessage[]>([]);

  // Folder loading callback, use useCallback to avoid re-creating
  const handleFoldersLoaded = React.useCallback((folders: MailFolder[]) => {
    setMailFolders(folders);
  }, []);

  // Handle messages loaded
  const handleMessagesLoaded = React.useCallback((newMessages: EmailMessage[]) => {
    setMessages(newMessages);
  }, []);

  React.useEffect(() => {
    if (!selectedTab) {
      const folder = mailFolders.find(folder => folder.displayName === 'Inbox');
      if (folder)
        setSelectedTab(folder.id);
    }
  }, [mailFolders, selectedTab]);

  // Clear messages when selected tab changes
  React.useEffect(() => {
    if (selectedTab) {
      setSelectedEmail(null);
      setMessages([]);
    }
  }, [selectedTab]); // This effect runs whenever selectedTab changes

  // Auto read selected email when selected email changes
  React.useEffect(() => {
    if (selectedEmail && !selectedEmail.isRead) {
      console.log('Marking email as read:', selectedEmail.id);
      readUnreadEmail(selectedEmail.id, true);
      setMailReadStatusInState(selectedEmail, true);
    }
  }, [selectedEmail]);
  
  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  // Mail folder tab component
  const MailFolder = (props: any) => {
    const folder = props.dataContext;

    return (
      <Tab
        key={folder.id}
        value={folder.id}
        className={styles.verticalTab}
      >
        {folder.displayName}
        {folder.unreadItemCount ? ` (${folder.unreadItemCount})` : ''}
      </Tab>
    );
  };

  const readUnreadEmail = async (emailId: string, isRead: boolean) => {
    try {
      const token = await getToken(msalInstance, account, ['Mail.Read']);
      const options = {
        authProvider: done => {
          done(null, token);
        }
      };
      const client = Client.init(options);
      await client.api(`/me/messages/${emailId}`)
        .patch({
          isRead: isRead
        });
    } catch (error) {
      console.error('Error marking email as read:', error);
    }
  }

  const setMailReadStatusInState = (email: EmailMessage, isRead: boolean) => {
    const updatedMessages = messages.map(message => {
      if (message.id === email.id) {
        return {
          ...message,
          isRead: isRead
        };
      }
      return message;
    });
    setMessages(updatedMessages);
    // reset selected email
    email.isRead = isRead;
  }

  // Email action buttons component
  const EmailActionButtons = () => (
    <div className={styles.buttonContainer}>
      <Button
        icon={<MailRead16Regular />}
        className={styles.readUnreadButton}
        // enable only when an email is selected
        {...(selectedEmail ? {} : { disabled: true })}
        onClick={() => {
          if (selectedEmail) {
            readUnreadEmail(selectedEmail.id, !selectedEmail.isRead);
            setMailReadStatusInState(selectedEmail, !selectedEmail.isRead);
          }
        }}
      >
        Read / Unread
      </Button>
    </div>
  );

  return (
    <>
      <PageHeader
        title={'Outlook'}
        description={<EmailActionButtons />}
      />

      {/* Use the standalone folder provider component */}
      <FolderProvider onFoldersLoaded={handleFoldersLoaded} />

      <div className={styles.verticalContainer}>
        {/* Folders Tab List - Left Side */}
        <TabList
          selectedValue={selectedTab}
          onTabSelect={onTabSelect}
          vertical
          className={styles.verticalTabList}
        >
          {mailFolders.length === 0 && (
            <Loading template="loading" message={`Loading mail folders...`}></Loading>
          )}
          {mailFolders.map(folder => (
            <MailFolder key={folder.id} dataContext={folder} />
          ))}
        </TabList>

        {/* Message List - Middle */}
        <div className={styles.messageList}>
          
        {selectedTab ? (
            <>
              {/* Messages provider to fetch data */}
              <MessageProvider
                folderId={selectedTab as string}
                onMessagesLoaded={handleMessagesLoaded}
              />

              {/* Render the messages from state */}
              {messages.length > 0 ? (
                messages.map(message => (
                  <div
                    key={message.id}
                    className={styles.messageItem}
                    onClick={() => setSelectedEmail(message)}>
                    <Messages email={message} />
                  </div>
                ))
              ) : (
                <div className={styles.emptyState}>No messages in this folder</div>
              )}
            </>
          ) : (
            <div className={styles.emptyState}>Select a folder to view messages</div>
          )}
        </div>

        {/* Content Area - Right Side */}
        <div className={styles.main}>
          {selectedEmail ? (
            <CustomGet
              key={`email-${selectedEmail.id}`}
              resource={`/me/messages/${selectedEmail.id}`}
              scopes={['Mail.Read']}
              pollingRate={0}
            >
              {(data, loading, error) => (
                <>
                  {data && <MailContent template="default" dataContext={data} />}
                  {loading && <Loading template="loading" message={`Loading message content...`} />}
                </>
              )}
            </CustomGet>
          ) : (
            <div className={styles.emptyState}>
              Select a message from the list to view its details
            </div>
          )}
        </div>
      </div>
    </>
  );
};