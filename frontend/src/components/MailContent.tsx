import { mergeClasses } from '@fluentui/react-components';
import { useEffect, useRef } from 'react';
import { useMailContentStyles } from '../styles/Styles';
import Person from '../MGTInternal/person/Person';
import { viewTypeConverter } from '../MGTInternal/components/viewTypeConverter';

export interface MgtTemplateProps<T = any> {
  template?: string;
  dataContext?: T;
}

export function MailContent(props: MgtTemplateProps) {
  const styles = useMailContentStyles();
  const email = props.dataContext;
  const contentRef = useRef<HTMLDivElement>(null);
  const senderAddress = email.sender?.emailAddress?.address || '';
  const senderName = email.sender?.emailAddress?.name || senderAddress || 'Unknown sender';

  // Format date and time
  const formattedDateTime = new Date(email.receivedDateTime).toLocaleString(undefined, {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });

  // Set HTML content safely
  useEffect(() => {
    if (contentRef.current && email.body?.content) {
      contentRef.current.innerHTML = email.body.content;
    }
  }, [email.body?.content]);

  return (
    <div className={styles.container}>
      {/* Header Card with Subject, Sender, Recipients and Time */}
      <div className={styles.headerCard}>
        {/* Subject */}
        <h1 className={styles.subject}>{email.subject || '(No subject)'}</h1>

        {/* Sender */}
        <div className={styles.senderSection}>
          <span className={styles.label}>From:</span>
          <Person
            personQuery={senderAddress}
            fallbackDetails={{
              id: senderAddress,
              displayName: senderName,
              mail: senderAddress,
            }}
            view={viewTypeConverter("oneline")}
            showPresence={true}
          />
        </div>

        {/* Recipients */}
        <div className={styles.recipientsSection}>
          <span className={styles.label}>To:</span>
          <div className={styles.recipientsList}>
            {email.toRecipients && email.toRecipients.map((recipient: any, index: number) => (
              <Person
                key={index}
                personQuery={recipient?.emailAddress?.address || ''}
                fallbackDetails={{
                  id: recipient?.emailAddress?.address || '',
                  displayName: recipient?.emailAddress?.name || recipient?.emailAddress?.address || 'Unknown recipient',
                  mail: recipient?.emailAddress?.address || '',
                }}
                view={viewTypeConverter("oneline")}
                showPresence={true}
              />
            ))}
          </div>
        </div>

        {/* CC Recipients if any */}
        {email.ccRecipients && email.ccRecipients.length > 0 && (
          <div className={styles.recipientsSection}>
            <span className={styles.label}>CC:</span>
            <div className={styles.recipientsList}>
              {email.ccRecipients.map((recipient: any, index: number) => (
                <Person
                  key={index}
                  personQuery={recipient?.emailAddress?.address || ''}
                  fallbackDetails={{
                    id: recipient?.emailAddress?.address || '',
                    displayName: recipient?.emailAddress?.name || recipient?.emailAddress?.address || 'Unknown recipient',
                    mail: recipient?.emailAddress?.address || '',
                  }}
                  view={viewTypeConverter("oneline")}
                  showPresence={true}
                />
              ))}
            </div>
          </div>
        )}

        {/* Date/Time as its own row at the bottom, aligned right */}
        <div className={styles.dateTimeContainer}>
          <span className={styles.dateTime}>{formattedDateTime}</span>
        </div>
      </div>

      {/* Email Content Card */}
      <div className={styles.contentCard}>
        {email.body?.content ? (
          <div
            ref={contentRef}
            className={styles.emailContent}
          />
        ) : (
          <div className={mergeClasses(styles.emailContent, styles.emptyContent)}>
            This email does not contain any content.
          </div>
        )}
      </div>
    </div>
  );
}