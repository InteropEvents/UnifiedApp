// import { Person } from '@microsoft/mgt-react';
import { mergeClasses } from '@fluentui/react-components';
import { personCardConverter } from '@microsoft/mgt-components/dist/es6/components/PersonCardInteraction';
import { useMessagesStyles } from '../styles/Styles';
import { MessageProps } from '../MGTInternal/components/MessageProps';
import { viewTypeConverter } from '../MGTInternal/components/viewTypeConverter';
import Person from '../MGTInternal/person/Person';


export function Messages({ email }: MessageProps) {
  const styles = useMessagesStyles();
  const isUnread = !email.isRead;
  const senderAddress = email.sender?.emailAddress?.address || '';
  const senderName = email.sender?.emailAddress?.name || senderAddress || 'Unknown sender';

  return (
    <label>
      <input
        type="radio"
        name="email"
        className={styles.seletedEmailRadio}
        aria-label={email.subject}
        onClick={(e) => {
          const currentLabel = e.currentTarget.closest('label');
          if (currentLabel) {
            const subjectElement = currentLabel.querySelector('#email-subject');
            if (subjectElement) {
              subjectElement.className = styles.subject;
            }
          }
        }}
      />
      <div className={styles.email}>
        <div className={styles.header}>
          <div>
            <Person
              personQuery={senderAddress}
              fallbackDetails={{
                id: senderAddress,
                displayName: senderName,
                mail: senderAddress,
              }}
              view={viewTypeConverter("oneline")}
              personCardInteraction={personCardConverter("hover")}
            />
          </div>
        </div>
        <div className={styles.title}>
          <h3
          id="email-subject"
           className={mergeClasses(
              styles.subject,
              isUnread && styles.unreadSubject
            )}>
            {email.subject}
          </h3>
          <span className={styles.date}>{new Date(email.receivedDateTime).toLocaleDateString()}</span>
        </div>
        {email.bodyPreview ?
          <div className={styles.body}>{email.bodyPreview}</div> :
          <div className={mergeClasses(styles.body, styles.emptyBody)}>...</div>}
      </div>
    </label>
  );
}
