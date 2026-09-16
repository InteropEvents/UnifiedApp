import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { Users } from '../components/Users';
import { ConditionalAccess } from '../components/ConditionalAccess';
import { useEntraIDStyles } from '../styles/Styles';
import { getGraphClient } from '../utils/graphClient';

export const EntraIDPage: React.FunctionComponent = () => {
  const styles = useEntraIDStyles();
  const [selectedNav, setSelectedNav] = React.useState('Users');

  return (
    <div className={styles.container}>
      {/* page header */}
      <PageHeader title={'EntraID'} description={'Manage your Azure Active Directory'} />

      {/* main */}
      <div className={styles.main}>
        {/* navigation */}
        <nav className={styles.nav}>
          <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
            <li
              className={`${styles.navItem} ${selectedNav === 'Users' ? 'active' : ''}`}
              onClick={() => setSelectedNav('Users')}
            >
              Users
            </li>
            <li
              className={`${styles.navItem} ${selectedNav === 'Conditional Access' ? 'active' : ''}`}
              onClick={() => setSelectedNav('Conditional Access')}
            >
              Conditional Access
            </li>
          </ul>
        </nav>

        {/* content */}
        <div style={{ flex: 1, padding: '1rem' }}>
          {selectedNav === 'Users' && <Users />}
          {selectedNav === 'Conditional Access' && <ConditionalAccess />}
        </div>
      </div>
    </div>
  );
};


/**
 * Check if the login user has enough permission to access this page
 * Checks if user has Global Administrator or Privileged Role Administrator role
 * @returns {Promise<boolean>} true if the user has permission, false otherwise
 */
export const hasEntraIDPermission = async (): Promise<boolean> => {
  try {
    const client = getGraphClient();

    const memberOfResponse = await client.api('/me/memberOf')
      .select('displayName')
      .get();

    const requiredRoles = ['Global Administrator', 'Privileged Role Administrator'];

    const hasRequiredRole = memberOfResponse.value?.some((item: any) =>
      requiredRoles.includes(item.displayName)
    );

    return hasRequiredRole || false;

  } catch (error) {
    console.error('Error checking Entra ID permissions:', error);
    return false;
  }
};