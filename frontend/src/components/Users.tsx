import * as React from 'react';
import {
  Button,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle,
  Input,
  Label,
  Checkbox,
  makeStyles
} from '@fluentui/react-components';
import { Providers, ProviderState } from '@microsoft/mgt';
import {useUserStyles} from '../styles/Styles';

interface UsersProps {}

export const Users: React.FunctionComponent<UsersProps> = () => {
  const styles = useUserStyles();
  const [users, setUsers] = React.useState<any[]>([]);
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [newUser, setNewUser] = React.useState<{
    accountEnabled: boolean;
    displayName: string;
    mailNickname: string;
    userPrincipalName: string;
    password: string;
    givenName: string | null;
    surname: string | null;
    userType: string;
  }>({
    accountEnabled: true,
    displayName: '',
    mailNickname: '',
    userPrincipalName: '',
    password: '',
    givenName: null,
    surname: null,
    userType: 'Member',
  });
  const [domain, setDomain] = React.useState<string>('');
  const [editingUserId, setEditingUserId] = React.useState<string | null>(null);
  const [deriveFromUPN, setDeriveFromUPN] = React.useState(true);
  const [filterText, setFilterText] = React.useState('');
  const [activeTab, setActiveTab] = React.useState<'basic' | 'properties' | 'assignments'>('basic');
  const [groups, setGroups] = React.useState<{ id: string; name: string }[]>([]);
  const [isGroupDialogOpen, setIsGroupDialogOpen] = React.useState(false); 
  const [selectedGroups, setSelectedGroups] = React.useState<{ id: string; name: string }[]>([]);
  const [groupFilterText, setGroupFilterText] = React.useState('');
  const [roles, setRoles] = React.useState<{ id: string; name: string }[]>([]);
  const [selectedRoles, setSelectedRoles] = React.useState<{ id: string; name: string }[]>([]);
  const [roleFilterText, setRoleFilterText] = React.useState('');
  const [isRoleDialogOpen, setIsRoleDialogOpen] = React.useState(false);
  const [initialGroups, setInitialGroups] = React.useState<{ id: string; name: string }[]>([]);
  const [initialRoles, setInitialRoles] = React.useState<{ id: string; name: string }[]>([]);
  const [hasPermission, setHasPermission] = React.useState<boolean | null>(null);

  React.useEffect(() => {
    const initializePage = async () => {
      const provider = Providers.globalProvider;
      if (provider && provider.state === ProviderState.SignedIn) {
        const graphClient = provider.graph.client;
  
        try {
          // check user permissions
          const permissionResponse = await graphClient.api('/me/memberOf').get();
          const hasPermission = permissionResponse.value.some((group: any) =>
            ['Global Administrator', 'User Administrator'].includes(group.displayName)
          );
          setHasPermission(hasPermission);
  
          if (!hasPermission) {
            return; // if no permission, exit early
          }
  
          // get default domain
          const domainResponse = await graphClient.api('/organization').get();
          const defaultDomain = domainResponse.value[0]?.verifiedDomains?.find((d: any) => d.isDefault)?.name;
          setDomain(defaultDomain || '');
  
          // get users
          const usersResponse = await graphClient.api('/users').get();
          const usersWithPhotos = await Promise.all(
            usersResponse.value.map(async (user: any) => {
              try {
                const photoResponse = await graphClient
                  .api(`/users/${user.id}/photo/$value`)
                  .responseType('blob' as any)
                  .get();
                const photoUrl = URL.createObjectURL(photoResponse);
                return { ...user, photoUrl };
              } catch (error) {
                return { ...user, photoUrl: null }; 
              }
            })
          );
          setUsers(usersWithPhotos);
        } catch (error) {
          console.error('Error initializing page:', error);
        }
      }
    };
  
    initializePage();
  }, []);

  if (hasPermission === null) {
    return <div>loading...</div>;
  }

  if (!hasPermission) {
    return <div>You don't have permission for this page</div>;
  }

  const handleSaveUser = async () => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;

    try {
      if (editingUserId) {
        // update user
        await graphClient.api(`/users/${editingUserId}`).patch({
          displayName: newUser.displayName,
          givenName: newUser.givenName,
          surname: newUser.surname,
          mailNickname: newUser.mailNickname,
          userPrincipalName: newUser.userPrincipalName,
          userType: newUser.userType,
        });
        assignUserToGroupsAndRoles(editingUserId);
        alert('User updated successfully!');
      } else {
        // create new user
        const user = {
          accountEnabled: newUser.accountEnabled,
          displayName: newUser.displayName,
          givenName: newUser.givenName,
          surname: newUser.surname,
          mailNickname: newUser.mailNickname,
          userPrincipalName: newUser.userPrincipalName,
          userType: newUser.userType,
          passwordProfile: {
            forceChangePasswordNextSignIn: true,
            password: newUser.password,
          },
        };
        const response = await graphClient.api('/users').post(user);
        const userId = response.id; // Get the ID of the newly created user
        assignUserToGroupsAndRoles(userId); // Assign groups and roles to the new user
        // Show success message
        alert('User added successfully!');
      }

      setIsDialogOpen(false);
      cleanEditingUser()
      setDeriveFromUPN(true); // Reset the checkbox state
      // refetch users after adding/updating
      const response = await graphClient.api('/users').get();
      setUsers(response.value);
    } catch (error) {
      console.error('Error saving user:', error);
      alert('Failed to save user. Please check the console for details.');
    }
  }
};
const assignUserToGroupsAndRoles = async (userId: string) => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;
    // Compare groups
    const groupsToAdd = selectedGroups.filter(
      (group) => !initialGroups.some((initialGroup) => initialGroup.id === group.id)
    );
    const groupsToRemove = initialGroups.filter(
      (initialGroup) => !selectedGroups.some((group) => group.id === initialGroup.id)
    );
    console.log('groupsToAdd:', groupsToAdd);
    console.log('groupsToRemove:', groupsToRemove);

    // Add groups
    for (const group of groupsToAdd) {
      try{
        await graphClient.api(`/groups/${group.id}/members/$ref`).post({
          "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`,
        });
      }catch (error) {
        alert('Error adding user to group: '+ error);
      }

    }

    // Remove groups
    for (const group of groupsToRemove) {
      try {
        await graphClient.api(`/groups/${group.id}/members/${userId}/$ref`).delete();
      } catch (error) {
        console.error(`Error removing user from group ${group.name}:`, error);
      }
    }

    // Compare roles
    const rolesToAdd = selectedRoles.filter(
      (role) => !initialRoles.some((initialRole) => initialRole.id === role.id)
    );
    const rolesToRemove = initialRoles.filter(
      (initialRole) => !selectedRoles.some((role) => role.id === initialRole.id)
    );
    console.log('rolesToAdd:', rolesToAdd);
    console.log('rolesToRemove:', rolesToRemove);

    // Add roles
    for (const role of rolesToAdd) {
      try{
        await graphClient.api('/roleManagement/directory/roleAssignments').post({
          principalId: userId, // The ID of the user
          roleDefinitionId: role.id, // The ID of the role definition
          directoryScopeId: '/', // Tenant-wide scope
        });
      }catch (error) {
        alert('Error adding user to role: '+ error);
      }

    }

    // Remove roles
    for (const role of rolesToRemove) {
      try {
        const response = await graphClient
          .api(`/roleManagement/directory/roleAssignments?$filter=principalId eq '${userId}' and roleDefinitionId eq '${role.id}'`)
          .get();

        if (response.value.length > 0) {
          const assignmentId = response.value[0].id;
          await graphClient.api(`/roleManagement/directory/roleAssignments/${assignmentId}`).delete();
        }
      } catch (error) {
        console.error(`Error removing user from role ${role.name}:`, error);
      }
    }
  }
};

const cleanEditingUser = () => {
    setEditingUserId(null); 
    setNewUser({
      accountEnabled:true, 
      displayName: '', 
      mailNickname: '', 
      userPrincipalName: '', 
      password: '',
      givenName: null,
      surname: null,
      userType: 'Member'
    });
    setSelectedGroups([]);
    setSelectedRoles([]);
    setActiveTab('basic'); // Reset to the first tab
}

const getInitials = (name: string) => {
  if (!name) return '';
  const nameParts = name.split(' ');
  const initials = nameParts.map((part) => part[0]).join('');
  return initials.toUpperCase().slice(0, 2); // Limit to 2 initials
};

const filteredUsers = users.filter((user) =>
  user.displayName.toLowerCase().includes(filterText.toLowerCase()) ||
  user.userPrincipalName.toLowerCase().includes(filterText.toLowerCase())
);

//fetch groups from graph api
const fetchGroups = async () => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;
    try {
      const response = await graphClient
      .api(
        '/groups?$select=displayName,mail,id,onPremisesSyncEnabled,onPremisesLastSyncDateTime,groupTypes,mailEnabled,securityEnabled,resourceProvisioningOptions,isAssignableToRole&$top=100'
      )
      .get();
      // filter out groups that are mailEnabled and securityEnabled
      const validGroups = response.value.filter(
        (group: any) => !(group.mailEnabled && group.securityEnabled)
      );
      setGroups(
        validGroups.map((group: any) => ({
          id: group.id,
          name: group.displayName,
        }))
      );
    } catch (error) {
      console.error('Error fetching groups:', error);
    }
  }
}
const filteredGroups = groups.filter((group) =>
  group.name.toLowerCase().includes(groupFilterText.toLowerCase())
);

const fetchRoles = async () => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;
    try {
      const response = await graphClient
        .api('/roleManagement/directory/roleDefinitions?$top=500')
        .version('beta')
        .get();
      // Filter roles that are assignable
      const validRoles = response.value.filter((role: any) => role.assignmentMode === 'allowed');

      setRoles(
        validRoles.map((role: any) => ({
          id: role.id,
          name: role.displayName,
        }))
      );
    } catch (error) {
      console.error('Error fetching roles:', error);
    }
  }
};
const fetchUserGroups = async (userId: string) => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;
    try {
      const response = await graphClient
        .api(`/users/${userId}/memberOf/microsoft.graph.group?$select=id,displayName`)
        .version('beta')
        .get();

      // Map the groups to the required format
      const userGroups = response.value.map((group: any) => ({
        id: group.id,
        name: group.displayName,
      }));

      // Update the selectedGroups state
      setSelectedGroups(userGroups);
      setInitialGroups(userGroups); // Save initial groups
    } catch (error) {
      console.error('Error fetching user groups:', error);
    }
  }
};
const fetchUserRoles = async (userId: string) => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;
    try {
      const response = await graphClient
        .api(`/roleManagement/directory/roleAssignments?$filter=principalId eq '${userId}'&$expand=roleDefinition`)
        .get();

      // Map the roles to the required format
      const userRoles = response.value.map((assignment: any) => ({
        id: assignment.roleDefinition.id,
        name: assignment.roleDefinition.displayName,
      }));

      // Update the selectedRoles state
      setSelectedRoles(userRoles);
      setInitialRoles(userRoles); // Save initial roles
    } catch (error) {
      console.error('Error fetching user roles:', error);
    }
  }
};

const fetchUserDetails = async (userId: string) => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;
    try {
      const response = await graphClient
        .api(`/users/${userId}`)
        .query({
          $select: 'accountEnabled,ageGroup,assignedLicenses,businessPhones,city,companyName,consentProvidedForMinor,country,createdDateTime,creationType,department,displayName,mail,employeeId,employeeHireDate,employeeOrgData,employeeType,onPremisesExtensionAttributes,externalUserStateChangeDateTime,faxNumber,givenName,imAddresses,identities,externalUserState,jobTitle,surname,lastPasswordChangeDateTime,legalAgeGroupClassification,mailNickname,mobilePhone,id,officeLocation,onPremisesSamAccountName,onPremisesDistinguishedName,onPremisesDomainName,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesProvisioningErrors,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,otherMails,passwordPolicies,passwordProfile,preferredDataLocation,preferredLanguage,proxyAddresses,signInSessionsValidFromDateTime,sponsors,state,streetAddress,usageLocation,userPrincipalName,userType,postalCode',
          $expand: 'manager',
        })
        .get();

      return response; // Return the user details
    } catch (error) {
      console.error('Error fetching user details:', error);
      return null;
    }
  }
  return null;
};


  return (
    <div>
      <h2>Users</h2>
      {/* Add User Button */}
      <div style={{ marginBottom: '1rem' }}>
        <Dialog open={isDialogOpen} onOpenChange={(event, data) => setIsDialogOpen(data.open)}>
          <DialogTrigger>
            <Button
                appearance="primary"
                onClick={() => {
                    cleanEditingUser()
                }}
                >
                Add User</Button>
          </DialogTrigger>
          <DialogSurface style={{ maxWidth: '750px' }}>
            <DialogBody className={styles.dialogBody} >
            <DialogTitle>{editingUserId ? 'Edit User' : 'Add New User'}</DialogTitle>
            {/* Tab Buttons */}
            <div style={{ display: 'flex', gap: '1rem', marginBottom: '1rem',borderBottom: '1px solid #ccc' }}>
              <Button
                className={styles.tabButton}
                onClick={() => setActiveTab('basic')}
              >
                Basic
              </Button>
              <Button
                className={styles.tabButton}
                onClick={() => setActiveTab('properties')}
              >
                Properties
              </Button>
              <Button
                className={styles.tabButton}
                onClick={() => setActiveTab('assignments')}
              >
                Assignments
              </Button>
            </div>

            <form 
                onSubmit={(e) => {
                    e.preventDefault();
                    handleSaveUser();
                }}
            >
                
                {activeTab === 'basic' && (
                <div className={styles.subcontainer}>
                <div className={styles.twoColumnRow}>
                <Label htmlFor="userPrincipalName" className={styles.label}>User Principal Name<span className={styles.required}>*</span></Label>
                <div style={{ display: 'flex', alignItems: 'center' }}>
                <Input
                    id="userPrincipalName"
                    value={newUser.userPrincipalName.split('@')[0]}
                    onChange={(e) => {
                        const prefix = e.target.value;
                        setNewUser({
                            ...newUser,
                            userPrincipalName: `${prefix}@${domain}`,
                            mailNickname: deriveFromUPN ? prefix : newUser.mailNickname, // Update mailNickname if deriveFromUPN is checked
                          });
                    }}
                    required
                    /><label>@{domain}</label></div>
                    </div>
                <div className={styles.twoColumnRow}>
                <Label htmlFor="mailNickname" className={styles.label}>Mail Nickname<span className={styles.required}>*</span></Label>
                <Input
                  id="mailNickname"
                  value={deriveFromUPN ? newUser.userPrincipalName.split('@')[0] : newUser.mailNickname}
                  onChange={(e,data) => {
                    setNewUser({ ...newUser, mailNickname: data.value });
                  }}
                  disabled={deriveFromUPN} //if deriveFromUPN is checked, disable this field
                  required
                />
                </div>
                <Checkbox
                  label="Derive from user principal name"
                  checked={deriveFromUPN}
                  onChange={(e, data) => {
                    setDeriveFromUPN(!!data.checked);
                  }}
                />
                <div className={styles.twoColumnRow}>
                <Label htmlFor="displayName" className={styles.label} >Display Name<span className={styles.required}>*</span></Label>
                <Input
                  id="displayName"
                  value={newUser.displayName}
                  onChange={(e) => setNewUser({ ...newUser, displayName: e.target.value })}
                  required
                />
                </div>
                {/* Only show password field in Add mode */}
                {!editingUserId && (
                  <div className={styles.twoColumnRow}>
                    <Label htmlFor="password" className={styles.label}>
                      Password<span className={styles.required}>*</span>
                    </Label>
                    <Input
                      id="password"
                      type="password"
                      value={newUser.password}
                      onChange={(e) => setNewUser({ ...newUser, password: e.target.value })}
                      required
                    />
                  </div>
                )}
                <Checkbox
                  label="Account Enabled"
                  checked={newUser.accountEnabled}
                  onChange={(e, data) => {
                    setNewUser({ ...newUser, accountEnabled: !!data.checked });
                  }}
                />
                </div>
                )}
                
                {activeTab === 'properties' && (
                  <div className={styles.subcontainer}>

                    {/* First Name */}
                <div className={styles.twoColumnRow}>
                  <Label htmlFor="firstName" className={styles.label}>
                    First Name
                  </Label>
                  <Input
                    id="firstName"
                    value={newUser.givenName || undefined}
                    onChange={(e) => setNewUser({ ...newUser, givenName: e.target.value })}
                  />
                </div>

                {/* Last Name */}
                <div className={styles.twoColumnRow}>
                  <Label htmlFor="lastName" className={styles.label}>
                    Last Name
                  </Label>
                  <Input
                    id="lastName"
                    value={newUser.surname || ''}
                    onChange={(e) => setNewUser({ ...newUser, surname: e.target.value })}
                  />
                </div>
                {/* User Type */}
                <div className={styles.twoColumnRow}>
                  <Label htmlFor="userType" className={styles.label}>
                    User Type
                  </Label>
                  <select
                    id="userType"
                    value={newUser.userType || 'Member'}
                    onChange={(e) => setNewUser({ ...newUser, userType: e.target.value })}
                    required
                    style={{
                      padding: '0.5rem',
                      border: '1px solid #ccc',
                      borderRadius: '4px',
                      width: '100%',
                    }}
                  >
                    <option value="Member">Member</option>
                    <option value="Guest">Guest</option>
                  </select>
                </div>               
                  </div>
                )}
                {activeTab === 'assignments' && (
                  <div className={styles.subcontainer}>
                    {/* Buttons */}
                    <div className={styles.subcontainer}>
                      <div style={{ display: 'flex', gap: '1rem', marginBottom: '1rem' }}>
                      <Button
                        appearance="secondary"
                        onClick={async () => {
                          try {
                            await fetchGroups();
                            setIsGroupDialogOpen(true); 
                          } catch (error) {
                            console.error('Error fetching groups:', error);
                          }
                        }}
                      >
                        Add Group
                      </Button>
                      <Button
                        appearance="secondary"
                        onClick={async () => {
                          try {
                            await fetchRoles();
                            setIsRoleDialogOpen(true);
                          } catch (error) {
                            console.error('Error fetching roles:', error);
                          }
                        }}
                      >
                        Add Role
                      </Button>
                      </div>
                      {selectedGroups.length == 0 && selectedRoles.length == 0 && (
                        <p style={{ color: 'gray' }}>No assignments</p>
                      )}
                      {(selectedGroups.length > 0 || selectedRoles.length > 0 ) && (
                        <div style={{ borderTop: '1px solid #ccc'}}>
                          <table className={styles.table}>
                            <thead>
                              <tr>
                                <th className={styles.th}>Type</th>
                                <th className={styles.th}>Name</th>
                              </tr>
                            </thead>
                            <tbody>
                              {selectedGroups.map((group) => (
                                <tr key={group.id}>
                                  <td className={styles.td}>Group</td>
                                  <td className={styles.td}>{group.name}</td>
                                </tr>
                              ))}
                                {selectedRoles.map((role) => (
                                  <tr key={role.id}>
                                    <td className={styles.td}>Role</td>
                                    <td className={styles.td}>{role.name}</td>
                                  </tr>
                                ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                    <Dialog open={isGroupDialogOpen} 
                    onOpenChange={(event, data) => {
                      setIsGroupDialogOpen(data.open);
                      setGroupFilterText(''); // Reset filter text when dialog opens
                      }}>
                      <DialogSurface style={{ maxWidth: '500px' }}>
                        <DialogBody className={styles.dialogBody}>
                          <DialogTitle>Select Groups</DialogTitle>
                            <Input
                            type="text"
                            placeholder="Filter groups..."
                            value={groupFilterText}
                            onChange={(e) => setGroupFilterText(e.target.value)}
                            className={styles.filterInput}
                          />
                            <ul className={styles.listContainer}>
                              {filteredGroups.map((group) => (
                                <li key={group.id}>
                                  <Checkbox
                                    label={group.name}
                                    checked={selectedGroups.some((selected) => selected.id === group.id)}
                                    onChange={(e, data) => {
                                      if (data.checked) {
                                        setSelectedGroups((prev) => [...prev, group]);
                                      } else {
                                        setSelectedGroups((prev) => prev.filter((selected) => selected.id !== group.id));
                                      }
                                    }}
                                  />
                                </li>
                              ))}
                            </ul>
                            <Button
                              appearance="primary"
                              className={styles.button}
                              onClick={() => {
                                console.log('Selected Groups:', selectedGroups);
                                setIsGroupDialogOpen(false); 
                              }}
                            >
                              Select
                            </Button>
                        </DialogBody>
                      </DialogSurface>
                    </Dialog>
                    <Dialog open={isRoleDialogOpen} 
                    onOpenChange={(event, data) => {
                      setIsRoleDialogOpen(data.open);
                      setRoleFilterText(''); // Reset filter text when dialog opens
                      }}>
                      <DialogSurface style={{ maxWidth: '500px' }}>
                        <DialogBody className={styles.dialogBody}>
                          <DialogTitle>Select Roles</DialogTitle>
                    
                          {/* Filter Box */}
                          <Input
                            type="text"
                            placeholder="Filter roles..."
                            value={roleFilterText}
                            onChange={(e) => setRoleFilterText(e.target.value)}
                            className={styles.filterInput}
                          />
                    
                          {/* Filtered Roles List */}
                          <ul className={styles.listContainer}>
                            {roles
                              .filter((role) => role.name.toLowerCase().includes(roleFilterText.toLowerCase()))
                              .map((role) => (
                                <li key={role.id}>
                                  <Checkbox
                                    label={role.name}
                                    checked={selectedRoles.some((selected) => selected.id === role.id)}
                                    onChange={(e, data) => {
                                      if (data.checked) {
                                        setSelectedRoles((prev) => [...prev, role]);
                                      } else {
                                        setSelectedRoles((prev) => prev.filter((selected) => selected.id !== role.id));
                                      }
                                    }}
                                  />
                                </li>
                              ))}
                          </ul>
                    
                          <Button
                              appearance="primary"
                              className={styles.button}
                              onClick={() => {
                                console.log('Selected Roles:', selectedRoles);
                                setIsRoleDialogOpen(false);
                              }}
                            >
                              Select
                            </Button>
                        </DialogBody>
                      </DialogSurface>
                    </Dialog>
                    
                  </div>
                )}
                <div style={{ justifyContent: 'center' }}>
                <Button type="submit" appearance="primary" >
                {editingUserId ? 'Save Changes' : 'Add User'}
                </Button>
                </div>
                
              </form>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>

      <Input
        type="text"
        placeholder="Filter users..."
        value={filterText}
        onChange={(e) => setFilterText(e.target.value)}
        className={styles.filterInput}
      />

      {/* User List */}
    <ul className={styles.userListContainer}>
      {filteredUsers.map((user) => (
        <li key={user.id} className={styles.listItem}>
          {user.photoUrl ? (
            <img
              src={user.photoUrl}
              alt={`${user.displayName}'s avatar`}
              className={styles.avatar}
            />
          ) : (
            <div className={styles.placeholderAvatar}>
              {getInitials(user.displayName)}
            </div>
          )}
          <div style={{ flex: 1 }}>
            <a
              href="#"
              className={styles.userLink}
              onClick={async (e) => {
                e.preventDefault();
                setEditingUserId(user.id); 
                console.log('Editing user :', user);
                // Fetch user details
                const userDetails = await fetchUserDetails(user.id);
                if (userDetails) {
                  setNewUser({
                    accountEnabled: userDetails.accountEnabled,
                    displayName: userDetails.displayName,
                    givenName: userDetails.givenName,
                    surname: userDetails.surname,
                    mailNickname: userDetails.mailNickname,
                    userPrincipalName: userDetails.userPrincipalName,
                    userType: userDetails.userType,
                    password: '', // Password is not returned for security reasons
                  });
            
                  // Optionally handle manager data
                  if (userDetails.manager) {
                    console.log('Manager:', userDetails.manager.displayName);
                  }
                }
                // Fetch and set the user's groups
                await fetchUserGroups(user.id);

                // Fetch and set the user's roles
                await fetchUserRoles(user.id);
                setIsDialogOpen(true);
              }}
            >
              <strong>{user.displayName}</strong>
            </a>
            <span> - {user.userPrincipalName}</span>
          </div>
        </li>
      ))}
    </ul>
    </div>
  );
};