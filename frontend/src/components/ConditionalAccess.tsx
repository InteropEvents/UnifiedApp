import * as React from 'react';
// import { Providers, ProviderState } from '@microsoft/mgt';
import { Providers, ProviderState } from '../MGTInternal/components/Providers';
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
import {Delete12Regular} from '@fluentui/react-icons';
import {useConditionalAccessStyles} from '../styles/Styles';

export const ConditionalAccess: React.FunctionComponent = () => {
  const styles = useConditionalAccessStyles();
  const [policies, setPolicies] = React.useState<any[]>([]); 
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [isUserDialogOpen, setIsUserDialogOpen] = React.useState(false); 
  const [isResourceDialogOpen, setIsResourceDialogOpen] = React.useState(false);
  const [newPolicy, setNewPolicy] = React.useState<{
    name: string;
    conditions: {
        clientAppTypes: string[],
        users: {
          includeUsers: string[];
        };
        applications: {
          includeApplications: string[];
        };
      };
    grantControls: {
        operator: 'AND' | 'OR';
        builtInControls: string[];
      };
    sessionControls: {
        applicationEnforcedRestrictions:{isEnabled:boolean} | null;
        cloudAppSecurity: {
            cloudAppSecurityType: 'mcasConfigured' | 'monitorOnly' | 'blockDownloads' | 'unknownFutureValue';
            isEnabled: boolean;
          } | null;
        disableResilienceDefaults: boolean | null;
        persistentBrowser:{
            isEnabled:boolean;
            mode: 'always' | 'never';
        } | null;
        signInFrequency:{
            isEnabled:boolean|null;
            value: number|null;
            type: 'hours' | 'days'|null;
            frequencyInterval:'timeBased'|'everyTime'|'unknownFutureValue'|null;
            authenticationType:'primaryAndSecondaryAuthentication'|'secondaryAuthentication'|'unknownFutureValue' |null;
        } | null;
    };
    }>({
    name: '',
    conditions: {
        clientAppTypes: ['mobileAppsAndDesktopClients', 'browser'],
        users: {
          includeUsers: [],
        },
        applications: {
          includeApplications: [],
        },
      },
      grantControls: {
        operator: 'OR', 
        builtInControls: [],
      }, 
    sessionControls: {
        applicationEnforcedRestrictions: null,
        cloudAppSecurity: null,
        disableResilienceDefaults: null,
        persistentBrowser:null,
        signInFrequency:null
    }
    });
  const [allUsers, setAllUsers] = React.useState<any[]>([]); // storage for all users
  const [selectedUsers, setSelectedUsers] = React.useState<string[]>([]); // storage for selected users
  const [allResources, setAllResources] = React.useState<any[]>([]); // storage for all resources
  const [selectedResources, setSelectedResources] = React.useState<string[]>([]); // storage for selected resources
  const [isGrantDialogOpen, setIsGrantDialogOpen] = React.useState(false); // control Grant Access dialog
  const [isSessionDialogOpen, setIsSessionDialogOpen] = React.useState(false); // control Session Controls dialog
  const [editingPolicyId, setEditingPolicyId] = React.useState<string | null>(null); // current editing policy ID
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [userFilter, setUserFilter] = React.useState('');
  const [resourceFilter, setResourceFilter] = React.useState('');
  const [hasPermission, setHasPermission] = React.useState<boolean | null>(null);

  const BUILT_IN_CONTROLS: readonly { value: string; label: string }[] = [
  { value: 'mfa', label: 'Require multifactor authentication' },
  { value: 'compliantDevice', label: 'Require device to be marked as compliant' },
  { value: 'domainJoinedDevice', label: 'Require Microsoft Entra hybrid joined device' },
  { value: 'approvedApplication', label: 'Require approved client app' },
  { value: 'compliantApplication', label: 'Require compliant application' },
  { value: 'passwordChange', label: 'Require password change' },
  { value: 'unknownFutureValue', label: 'Unknown future value' },
];

const SESSION_CONTROLS = [
  { value: 'applicationEnforcedRestrictions', label: 'Use app enforced restrictions' ,desc:'Session control to enforce application restrictions. Only Exchange Online and Sharepoint Online support this session control.'},
  { value: 'signInFrequency', label: 'Sign-in frequency' },
  { value: 'persistentBrowser', label: 'Persistent browser session' },
  { value: 'disableResilienceDefaults', label: 'Disable resilience defaults' },
  { value: 'cloudAppSecurity', label: 'cloud App Security' },
];
React.useEffect(() => {
  const initializePage = async () => {
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      const graphClient = provider.graph.client;

      try {
        // check if user has permission to access this page
        const permissionResponse = await graphClient.api('/me/memberOf').get();
        const hasPermission = permissionResponse.value.some((group: any) =>
          ['Global Administrator', 'Conditional Access Administrator'].includes(group.displayName)
        );
        setHasPermission(hasPermission);

        if (!hasPermission) {
          return; 
        }

        // get all users
        const usersResponse = await graphClient.api('/users').get();
        setAllUsers(usersResponse.value);

        // get all conditional access policies
        const policiesResponse = await graphClient.api('/identity/conditionalAccess/policies').get();
        setPolicies(policiesResponse.value);

        // get all service principals and applications
        const servicePrincipalsResponse = await graphClient.api('/servicePrincipals').get();
        const servicePrincipals = servicePrincipalsResponse.value.map((sp) => ({
          id: sp.id,
          appId: sp.appId,
          displayName: sp.displayName || 'Unnamed Service Principal',
          type: 'servicePrincipal',
        }));

        const applicationsResponse = await graphClient.api('/applications').get();
        const applications = applicationsResponse.value.map((app) => ({
          id: app.id,
          appId: app.appId,
          displayName: app.displayName || 'Unnamed Application',
          type: 'application',
        }));

        const combinedResources = [...servicePrincipals, ...applications];
        const uniqueResources = combinedResources.filter(
          (resource, index, self) =>
            index === self.findIndex((r) => r.appId === resource.appId)
        );

        setAllResources(uniqueResources);
      } catch (error) {
        console.error('Error initializing page:', error);
      }
    }
  };

  initializePage();
}, []);

    if (hasPermission === null) {
      return <div>Loading...</div>; 
    }
    
    if (!hasPermission) {
      return <div>You don't have permission for this page</div>; 
    }
  const handleUserSelection = (userId: string, isChecked: boolean) => {
    if (isChecked) {
      setSelectedUsers((prev) => [...prev, userId]);
    } else {
      setSelectedUsers((prev) => prev.filter((id) => id !== userId));
    }
  };

  const handleResourceSelection = (resourceId: string, isChecked: boolean) => {
    if (isChecked) {
      setSelectedResources((prev) => [...prev, resourceId]);
    } else {
      setSelectedResources((prev) => prev.filter((appId) => appId !== resourceId));
    }
  };

  // Select All checkbox handler
  const handleSelectAll = (isChecked: boolean, type: 'users' | 'resources') => {
    if (type === 'users') {
      if (isChecked) {
        setSelectedUsers(allUsers.map((user) => user.id));
      } else {
        setSelectedUsers([]);
      }
    } else if (type === 'resources') {
      if (isChecked) {
        setSelectedResources(allResources.map((resource) => resource.appId));
      } else {
        setSelectedResources([]);
      }
    }
  };

  const handleConfirmUserSelection = () => {
    setIsUserDialogOpen(false);
  };


  const handleConfirmResourceSelection = () => {
    setIsResourceDialogOpen(false); 
  };

  const handleEditPolicy = async (policyId: string) => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;

    try {
      const response = await graphClient.api(`/identity/conditionalAccess/policies/${policyId}`).get();
      console.log('Policy Details:', response);
      setSelectedUsers(response.conditions?.users?.includeUsers || []); 
      setSelectedResources(response.conditions?.applications?.includeApplications || []); 
      setNewPolicy({
        name: response.displayName,
        conditions: {
            clientAppTypes: response.conditions?.clientAppTypes || ['mobileAppsAndDesktopClients', 'browser'],
            users: {
              includeUsers: response.conditions?.users?.includeUsers || [],
            },
            applications: {
              includeApplications: response.conditions?.applications?.includeApplications || [],
            },
          },
        grantControls: {
          operator: response.grantControls?.operator || 'OR',
          builtInControls: response.grantControls?.builtInControls || [],
        },
        sessionControls: Object.keys(response.sessionControls || {}).reduce(
            (acc, key) => {
              const control = response.sessionControls[key];
              acc[key] = control !== null ? control : null; 
              return acc;
            },
            {
              applicationEnforcedRestrictions: null,
              cloudAppSecurity: null,
              disableResilienceDefaults: null,
              persistentBrowser: null,
              signInFrequency: null,
            } 
          ),
      });
      setEditingPolicyId(policyId); 
      setIsDialogOpen(true); 
    } catch (error) {
      console.error('Error fetching policy details:', error);
    }
  }
};

const handleSavePolicy = async () => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;

    try {
      setIsSubmitting(true);
      const newPolicyData = {
        displayName: newPolicy.name,
        state: 'enabled',
        conditions: {
            clientAppTypes: ['mobileAppsAndDesktopClients', 'browser'],
          users: {
            includeUsers: selectedUsers.length > 0 ? selectedUsers : [],
          },
          applications: {
            includeApplications: selectedResources.length > 0 ? selectedResources : [],
          },
        },
        grantControls: newPolicy.grantControls?.builtInControls?.length > 0
          ? {
              operator: newPolicy.grantControls.operator,
              builtInControls: newPolicy.grantControls.builtInControls,
            }
          : undefined,
        sessionControls: Object.keys(newPolicy.sessionControls).length > 0 
          ? newPolicy.sessionControls
          : undefined,
      };

      if (editingPolicyId) {
        // updaete existing policy
        await graphClient.api(`/identity/conditionalAccess/policies/${editingPolicyId}`).patch(newPolicyData);
        alert('Policy updated successfully!');
      } else {
        // create new policy
        await graphClient.api('/identity/conditionalAccess/policies').post(newPolicyData);
        alert('Policy created successfully!');
      }
       // refresh policies list after saving
       setTimeout(async () => {
        const response = await graphClient.api('/identity/conditionalAccess/policies').query({ timestamp: Date.now() }).get();
        setPolicies(response.value);
      }, 1000); // wait 1 second
      setIsDialogOpen(false);
      cleanEditingPolicy()
    } catch (error) {
      console.error('Error saving policy:', error);
      alert('Failed to save policy:' + error);
    } finally{
        setIsSubmitting(false); //reset submitting state
    }
  }
};

const cleanEditingPolicy = () => {
    setEditingPolicyId(null);
    setNewPolicy({
        name: '',
        conditions: {
            clientAppTypes: ['mobileAppsAndDesktopClients', 'browser'],
            users: {
              includeUsers: [],
            },
            applications: {
              includeApplications: [],
            },
          },
          grantControls: {
            operator: 'OR',
            builtInControls: [],
          },
          sessionControls: {
            applicationEnforcedRestrictions: null,
            cloudAppSecurity: null,
            disableResilienceDefaults: null,
            persistentBrowser:null,
            signInFrequency:null
        }
      });
      setSelectedUsers( []);
      setSelectedResources([]);
}

const handleDeletePolicy = async (policyId: string) => {
  const provider = Providers.globalProvider;
  if (provider && provider.state === ProviderState.SignedIn) {
    const graphClient = provider.graph.client;

    const confirmDelete = window.confirm('Are you sure you want to delete this policy?');
    if (!confirmDelete) {
      return; // if user cancels, do nothing
    }

    try {
      await graphClient.api(`/identity/conditionalAccess/policies/${policyId}`).delete();
      alert('Policy deleted successfully!');
      // refresh policies list after deleting
      setTimeout(async () => {
        const response = await graphClient.api('/identity/conditionalAccess/policies').query({ timestamp: Date.now() }).get();
        setPolicies(response.value);
      }, 1000); // wait 1 second
    } catch (error) {
      console.error('Error deleting policy:', error);
      alert('Failed to delete policy: ' + error);
    }
  }
};
const getSessionControlByValue = (value: string) => {
    return SESSION_CONTROLS.find((control) => control.value === value) || null;
  };
const setSignInFrequency = (key: keyof NonNullable<typeof newPolicy.sessionControls.signInFrequency>, value: any) => {
    console.log('Setting signInFrequency:', key, value);
  setNewPolicy({
    ...newPolicy,
    sessionControls: {
      ...newPolicy.sessionControls,
      signInFrequency: {
        ...newPolicy.sessionControls.signInFrequency,
        isEnabled: newPolicy.sessionControls.signInFrequency?.isEnabled || false,
        value: key === 'value' ? value : newPolicy.sessionControls.signInFrequency?.value || 0,
        type: key === 'type' ? value : newPolicy.sessionControls.signInFrequency?.type || 'hours',
        frequencyInterval: key === 'frequencyInterval' ? value : newPolicy.sessionControls.signInFrequency?.frequencyInterval || 'timeBased',
        authenticationType: key === 'authenticationType' ? value : newPolicy.sessionControls.signInFrequency?.authenticationType || 'primaryAndSecondaryAuthentication',
      },
    },
  });
};
const filteredUsers = allUsers.filter((user) =>
  user.displayName.toLowerCase().includes(userFilter.toLowerCase())
);
const filteredResources = allResources.filter((resource) =>
  resource.displayName.toLowerCase().includes(resourceFilter.toLowerCase())
);

  return (
    <div>
      <h2 className={styles.sectionHeader} >Conditional Access Policies</h2>
      {/* create  policy button and dialog*/}
      <Dialog open={isDialogOpen} onOpenChange={(event, data) => setIsDialogOpen(data.open)}>
         <DialogTrigger>
          <Button
            appearance="primary"
            onClick={() => {
                cleanEditingPolicy()
            }}
          >
            Create New Policy
          </Button>
        </DialogTrigger>
        <DialogSurface>
          <DialogBody className={styles.dialogBody}>
          <DialogTitle className={styles.sectionHeader}>{editingPolicyId ? 'Edit Policy' : 'Create New Policy'}</DialogTitle>
            <form
              onSubmit={(e) => {
                e.preventDefault();

                // check if policy name is empty
                if (!newPolicy.name.trim()) {
                    alert('Policy Name is required.');
                    return;
                }
                handleSavePolicy();
                
              }}
            >
              <Label htmlFor="policyName">
                Policy Name <span className={styles.required}>*</span>
              </Label>
              <Input
                id="policyName"
                value={newPolicy.name}
                onChange={(e) => setNewPolicy({ ...newPolicy, name: e.target.value })}
                required
              />
              {/* Assignments */}
              <h4 className={styles.sectionHeader}>Assignments</h4>
              <div className={styles.subcontainer}>
                <div className={styles.subcontainer}>
                <Label>Users</Label>
                <a
                href="#"
                className={styles.linkButton}
                onClick={(e) => {
                    e.preventDefault();
                    setIsUserDialogOpen(true);
                }}
                >
                {selectedUsers.length} users selected
                </a>
                {/* users dialogdialog */}
                <Dialog open={isUserDialogOpen} 
                onOpenChange={(event, data) => {
                  setIsUserDialogOpen(data.open);
                  setUserFilter(''); // reset filter when dialog opens
                  }}>
                    <DialogSurface>
                    <DialogBody className={styles.dialogBody}>
                        <DialogTitle>Select Users</DialogTitle>
                        <Input
                          type="text"
                          placeholder="Filter users..."
                          value={userFilter}
                          onChange={(e) => setUserFilter(e.target.value)}
                          style={{ marginBottom: '1rem', width: '100%' }}
                        />
                        <Checkbox
                        label="Select All"
                        checked={selectedUsers.length === allUsers.length && allUsers.length > 0}
                        onChange={(e, data) => handleSelectAll(!!data.checked, 'users')}
                        />
                        <ul className={styles.listContainer}>
                        {filteredUsers.map((user) => (
                            <li key={user.id}>
                            <Checkbox
                                label={user.displayName}
                                checked={selectedUsers.includes(user.id)}
                                onChange={(e, data) => handleUserSelection(user.id, !!data.checked)}
                            />
                            </li>
                        ))}
                        </ul>
                        <Button appearance="primary" onClick={handleConfirmUserSelection} className={styles.button}>
                        Select
                        </Button>
                    </DialogBody>
                    </DialogSurface>
                </Dialog>
                </div>
            <div className={styles.subcontainer}>
              <Label>Target Resources</Label>
              <a
                href="#"
                className={styles.linkButton} 
                onClick={(e) => {
                    e.preventDefault(); 
                    setIsResourceDialogOpen(true);
                }}
                >
                {selectedResources.length > 0
                    ? `${selectedResources.length} resources selected`
                    : 'No target resources selected'}
                </a>

              {/* Target resource dialog*/}
              <Dialog open={isResourceDialogOpen}
               onOpenChange={(event, data) => {
                setIsResourceDialogOpen(data.open);
                setResourceFilter(''); // reset filter when dialog opens
                }}>
                <DialogSurface>
                  <DialogBody className={styles.dialogBody}>
                    <DialogTitle>Select Resources</DialogTitle>
                    <Input
                      type="text"
                      placeholder="Filter resources..."
                      value={resourceFilter}
                      onChange={(e) => setResourceFilter(e.target.value)}
                      style={{ marginBottom: '1rem', width: '100%' }}
                    />
                    <Checkbox
                      label="Select All"
                      checked={selectedResources.length === allResources.length && allResources.length > 0}
                      onChange={(e, data) => handleSelectAll(!!data.checked, 'resources')}
                    />
                    <ul className={styles.listContainer}>
                      {filteredResources.map((resource) => (
                        <li key={resource.appId}>
                          <Checkbox
                            label={resource.displayName}
                            checked={selectedResources.includes(resource.appId)}
                            onChange={(e, data) => handleResourceSelection(resource.appId, !!data.checked)}
                          />
                        </li>
                      ))}
                    </ul>
                    <Button className={styles.button} appearance="primary" onClick={handleConfirmResourceSelection}>
                      Select
                    </Button>
                  </DialogBody>
                </DialogSurface>
              </Dialog>
              </div>
              </div>
            <h4 className={styles.sectionHeader}>Access Controls</h4>
            <div className={styles.subcontainer}>
            <div className={styles.subcontainer}>
            <Label>Grant</Label>
            <a
              href="#"
              className={styles.linkButton}
              onClick={(e) => {
                e.preventDefault();
                setIsGrantDialogOpen(true); 
              }}
            >
              {newPolicy.grantControls?.builtInControls?.length > 0
                ? `${newPolicy.grantControls.builtInControls.length} controls selected`
                : '0 controls selected'}
            </a>
            
            {/* Grant Access dialog */}
            <Dialog open={isGrantDialogOpen} onOpenChange={(event, data) => setIsGrantDialogOpen(data.open)}>
              <DialogSurface>
              <DialogBody className={styles.dialogBody}>
                <DialogTitle>Grant Access</DialogTitle>
                {BUILT_IN_CONTROLS.map((control) => (
                    <Checkbox
                    key={control.value}
                    label={control.label}
                    checked={newPolicy.grantControls.builtInControls.includes(control.value)}
                    onChange={(e, data) => {
                        const updatedControls = data.checked
                        ? [...newPolicy.grantControls.builtInControls, control.value]
                        : newPolicy.grantControls.builtInControls.filter((c) => c !== control.value);
                        setNewPolicy({
                            ...newPolicy,
                            grantControls: {
                              ...newPolicy.grantControls, 
                              builtInControls: updatedControls,
                            },
                          });
                    }}
                    />
                ))}
                {/* Radio Button Group for Operator */}
                <div style={{ marginTop: '1rem' }}>
                    <Label>For multiple controls:</Label>
                    <div style={{ display: 'flex', gap: '1rem', marginTop: '0.5rem' }}>
                    <label>
                        <input
                        type="radio"
                        name="grantOperator"
                        value="AND"
                        checked={newPolicy.grantControls.operator === 'AND'}
                        onChange={() =>
                            setNewPolicy({
                              ...newPolicy,
                              grantControls: {
                                ...newPolicy.grantControls,
                                operator: 'AND', // reset operator to AND
                              },
                            })
                          }
                        />
                        Require all the selected controls
                    </label>
                    <label>
                        <input
                        type="radio"
                        name="grantOperator"
                        value="OR"
                        checked={newPolicy.grantControls.operator === 'OR'}
                        onChange={() =>
                            setNewPolicy({
                              ...newPolicy,
                              grantControls: {
                                ...newPolicy.grantControls,
                                operator: 'OR', // reset operator to OR
                              },
                            })
                          }
                        />
                        Require one of the selected controls
                    </label>
                    </div>
                </div>                
                <Button appearance="primary" onClick={() => setIsGrantDialogOpen(false)}>
                    Select
                </Button>
                </DialogBody>
              </DialogSurface>
            </Dialog>
            </div>
            <div className={styles.subcontainer}>
            <Label>Session</Label>
            <a
              href="#"
              className={styles.linkButton} 
              onClick={(e) => {
                e.preventDefault(); 
                setIsSessionDialogOpen(true);
              }}
            >
              {Object.values(newPolicy.sessionControls).filter((control) => control !== null).length > 0
            ? `${Object.values(newPolicy.sessionControls).filter((control) => control !== null).length} controls selected`
            : '0 controls selected'}
            </a>

            {/* Session Controls dialog */}
            <Dialog open={isSessionDialogOpen} onOpenChange={(event, data) => setIsSessionDialogOpen(data.open)}>
            <DialogSurface>
                <DialogBody className={styles.dialogBody}>
                  <DialogTitle>Session Controls</DialogTitle>
                  <Checkbox
                  label={getSessionControlByValue("applicationEnforcedRestrictions")?.label}
                  checked={newPolicy.sessionControls.applicationEnforcedRestrictions?.isEnabled || false}
                  onChange={(e, data) => {
                    setNewPolicy({
                      ...newPolicy,
                      sessionControls: {
                        ...newPolicy.sessionControls,
                        applicationEnforcedRestrictions: data.checked
                          ? { isEnabled: true } 
                          : null, 
                      },
                    });
                  }}
                />
                {newPolicy.sessionControls.applicationEnforcedRestrictions?.isEnabled && (
                    <p style={{ marginLeft: '1.5rem', fontSize: '0.9rem', color: 'gray',marginTop:'0',backgroundColor:'azure' }}>
                        This control only works with supported apps. Currently, Office 365, Exchange Online, and SharePoint Online are
                        the only cloud apps that support app enforced restrictions.{' '}
                        <a href="https://learn.microsoft.com" target="_blank" rel="noopener noreferrer">
                        Learn more
                        </a>
                    </p>
                 )}
                <Checkbox
                    label={getSessionControlByValue("signInFrequency")?.label}
                    checked={newPolicy.sessionControls.signInFrequency?.isEnabled || false}
                    onChange={(e, data) => 
                        setNewPolicy({
                            ...newPolicy,
                            sessionControls: {
                              ...newPolicy.sessionControls,
                              signInFrequency: data.checked
                              ?{
                                ...newPolicy.sessionControls.signInFrequency,
                                isEnabled: !!data.checked,
                                frequencyInterval: null,
                                value:  null,
                                type:  null,
                                authenticationType: null,
                              }:null
                            },
                          })
                    }
                />
                {newPolicy.sessionControls.signInFrequency?.isEnabled && (
                    <div style={{ marginLeft: '1.5rem'}}>
                        {/* Radio Button: Periodic reauthentication */}
                        <label>
                        <input
                            type="radio"
                            name="signInFrequency"
                            value="timeBased"
                            checked={newPolicy.sessionControls.signInFrequency.frequencyInterval === 'timeBased' ||
                                newPolicy.sessionControls.signInFrequency.frequencyInterval === null}
                            onChange={() =>setSignInFrequency('frequencyInterval','timeBased')}
                        />
                        Periodic reauthentication
                        </label>

                        {/*input:  */}
                        {(newPolicy.sessionControls.signInFrequency.frequencyInterval === null ||
                        newPolicy.sessionControls.signInFrequency.frequencyInterval === 'timeBased') && (
                        <div style={{ marginLeft: '1.5rem', marginTop: '0.5rem' }}>
                            <div style={{ gap: '0.5rem' }}  className={styles.subcontainer}>
                            <Input
                              type="number"
                              value={newPolicy.sessionControls.signInFrequency.value?.toString() || '0'} //
                              onChange={(e) => {
                                const inputValue = parseInt(e.target.value, 10); 
                                const type = newPolicy.sessionControls.signInFrequency?.type;
                            
                                // 验证范围
                                if (type && type === 'hours' && (inputValue < 1 || inputValue > 23)) {
                                  alert('For "Hours", the value must be between 1 and 23.');
                                  return;
                                }
                                if (type && type === 'days' && (inputValue < 1 || inputValue > 365)) {
                                  alert('For "Days", the value must be between 1 and 365.');
                                  return;
                                }
                            
                                setSignInFrequency('value', inputValue || 0);
                              }}
                              style={{ width: '4rem' }}
                            />
                            <select style={{ width: '6rem',height:'2rem' }}
                                value={newPolicy.sessionControls.signInFrequency.type || 'hours'}
                                onChange={(e) => setSignInFrequency('type', e.target.value as 'hours' | 'days')}
                            >
                                <option value="hours">Hours</option>
                                <option value="days">Days</option>
                            </select>
                            </div>
                        </div>
                        )}

                        {/* Radio Button: Every time */}
                        <label style={{ marginTop: '1rem', display: 'block' }}>
                        <input
                            type="radio"
                            name="signInFrequency"
                            value="everyTime"
                            checked={newPolicy.sessionControls.signInFrequency.frequencyInterval === 'everyTime'}
                            onChange={() =>
                            setNewPolicy({
                                ...newPolicy,
                                sessionControls: {
                                ...newPolicy.sessionControls,
                                signInFrequency: {
                                    ...newPolicy.sessionControls.signInFrequency,
                                    frequencyInterval: 'everyTime',
                                    value: null,
                                    type: null,
                                    isEnabled: true,
                                    authenticationType: 'primaryAndSecondaryAuthentication',
                                },
                                },
                            })
                            }
                        />
                        Every time
                        </label>
                    </div>
                    )}
                <Checkbox
                    label={getSessionControlByValue('disableResilienceDefaults')?.label}
                    checked={newPolicy.sessionControls.disableResilienceDefaults || false}
                    onChange={(e, data) => {
                        setNewPolicy({
                          ...newPolicy,
                          sessionControls: {
                            ...newPolicy.sessionControls,
                            disableResilienceDefaults: !!data.checked,
                          },
                        });
                      }}
                />

                  <Button appearance="primary" onClick={() => setIsSessionDialogOpen(false)}>
                    Select
                  </Button>
                </DialogBody>
            </DialogSurface>
            </Dialog>
            </div>
            </div>
            <Button
              type="submit"
              appearance="primary"
              className={styles.submitButton}
              disabled={isSubmitting}
            >
              {isSubmitting ? 'Processing...' : editingPolicyId ? 'Save Changes' : 'Create'}
            </Button>
            </form>
          </DialogBody>
        </DialogSurface>
      </Dialog>

    <table className={styles.table}>
      <thead>
        <tr>
          <th className={styles.th}>Policy Name</th>
          <th className={styles.th}>State</th>
          <th className={styles.th}>Created</th>
          <th className={styles.th}>Actions</th>
        </tr>
      </thead>
      <tbody>
        {policies.map((policy) => (
          <tr key={policy.id}>
            <td className={styles.td}>
              <a
                href="#"
                className={styles.link}
                onClick={(e) => {
                  e.preventDefault();
                  handleEditPolicy(policy.id);
                }}
              >
                {policy.displayName}
              </a>
            </td>
            <td className={styles.td}>{policy.state}</td>
            <td className={styles.td}>{new Date(policy.createdDateTime).toLocaleDateString()}</td>
            <td className={styles.td}>
                <button
                    className={styles.deleteButton}
                    onClick={() => handleDeletePolicy(policy.id)}
                >
                    <Delete12Regular className={styles.deleteIcon} /> {/* icon */}
                    Delete
                </button>
                </td>
          </tr>
        ))}
      </tbody>
    </table>
    </div>
  );
};