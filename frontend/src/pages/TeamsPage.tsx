import * as React from 'react';
import { ResponseType, Client } from '@microsoft/microsoft-graph-client';
// Step1: The equivalent of mgt-file in typescript is FileList. 
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { CustomFile } from '../components/CustomFile';
import { Button } from '@fluentui/react-components';
import { Tree, TreeItem, TreeItemLayout } from '@fluentui/react-tree';
import { ChevronRightRegular } from '@fluentui/react-icons';
import { Team, Channel } from '@microsoft/microsoft-graph-types';
import { PageHeader } from '../components/PageHeader';
import { Loading } from '../components/Loading';
import { useTeamsPageStyles } from '../styles/Styles'; // Import the common styles
import { FileList } from '../MGTInternal/FileList/FileList';
import '../MGTInternal/FileList/FileList.css';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig } from '../MGTInternal/login/msalConfig';

const msalInstance = new PublicClientApplication(msalConfig);

const ChannelsTree = (props) => {
    const [channels, setChannels] = React.useState<Channel[]>([]);
    const [selectedChannelId, setSelectedChannelId] = React.useState('');
    const {
        teamId,
        graphClient,
        selectedTeamId,
        setSelectedTeamId,
        setSelectedChannelName
    } = props;
    // Step3: getSelectedTeamChannel, we firstset TeamId, Channel Name and ChannelId.
    // Finally, we call GraphAPI to list all the children of a driveitem which will 
    // fetch all the files in a given channel. 
    const getSelectedTeamChannel = (channelName, channelId) => {
        setSelectedTeamId(teamId);
        setSelectedChannelName(channelName);
        setSelectedChannelId(channelId);

        let apiCon = [{
            api: `https://graph.microsoft.com/v1.0/groups/${teamId}/drive/root:/${channelName}:/children`,
            type: "GET"
        }];
        PubSub.publish("Calendar", apiCon);
    }

    React.useEffect(() => {
        let cancelled = false;

        const loadChannels = async () => {
            try {
                const currentChannels = await getChannelsByTeam(graphClient, teamId);
                if (!cancelled) setChannels(currentChannels);
            } catch (error) {
                console.error(error);
            }
        };

        loadChannels();
        return () => {
            cancelled = true;
        };
    }, [graphClient, teamId]);

    return (
        <Tree aria-label='ChannelTree'>
            {channels.map((channel) => (
                <TreeItem
                    itemType='leaf'
                    key={channel.id}
                    onClick={() => {
                        getSelectedTeamChannel(channel.displayName, channel.id)
                    }}
                >
                    <TreeItemLayout
                        aside={(selectedChannelId === channel.id && selectedTeamId === teamId) ?
                            <ChevronRightRegular /> : null}
                    >
                        <div style={{ marginLeft: '1.2vh' }}>
                            {channel.displayName}
                        </div>
                    </TreeItemLayout>
                </TreeItem>
            ))}
        </Tree>
    );
}

const TeamImg = (props) => {
    const { graphClient, teamId } = props;
    const [teamPhoto, setTeamPhoto] = React.useState('');

    React.useEffect(() => {
        let cancelled = false;

        const loadPhoto = async () => {
            try {
                const photo = await getTeamPhoto(graphClient, teamId) as string;
                if (!cancelled) setTeamPhoto(photo);
            } catch (error) {
                console.error(error);
            }
        };

        loadPhoto();
        return () => {
            cancelled = true;
        };
    }, [graphClient, teamId]);

    return (
        <img
            src={teamPhoto}
            alt=""
            aria-hidden="true"
            style={{ width: '25px', borderRadius: '4px', marginRight: '1vh' }}
        />
    );
}

// Step2: The Graph toolkit does not currently offer any existing components for displaying Teams/Channel views. 
// Therefore, we developed our own custom component , ChannelFilesPage, to provide this functionality
//main component
export const ChannelFilesPage: React.FunctionComponent = () => {
    const [loading, setLoading] = React.useState(false);
    const [teams, setTeams] = React.useState<Team[]>([]);
    const [selectedTeamId, setSelectedTeamId] = React.useState('');
    const [selectedChannelName, setSelectedChannelName] = React.useState('');
    const [isPreviewVisible, setIsPreviewVisible] = React.useState(false);
    const [previewUrl, setPreviewUrl] = React.useState<string>('');
    const [accessToken, setAccessToken] = React.useState('');
    const [fileListLoading, setFileListLoading] = React.useState(true);
    const styles = useTeamsPageStyles();
    const fileListRef = React.useRef<HTMLDivElement>(null);

    // 使用 msalInstance 获取 accessToken
    React.useEffect(() => {
        const fetchToken = async () => {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                try {
                    const response = await msalInstance.acquireTokenSilent({
                        account: accounts[0],
                        scopes: ['User.Read', 'Group.Read.All', 'Channel.ReadBasic.All']
                    });
                    setAccessToken(response.accessToken);
                } catch (error) {
                    console.error('获取 accessToken 失败:', error);
                }
            }
        };
        fetchToken();
    }, []);

    // 用 accessToken 创建 graphClient
    const graphClient = React.useMemo(() => {
        if (!accessToken) return null;
        return Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
    }, [accessToken]);

    React.useEffect(() => {
        if (!graphClient) return;

        let cancelled = false;
        const loadTeams = async () => {
            setLoading(true);
            try {
                const currentTeams = await getAllMyTeams(graphClient);
                if (!cancelled) setTeams(currentTeams);
            } catch (error) {
                console.error(error);
            } finally {
                if (!cancelled) setLoading(false);
            }
        };

        loadTeams();
        return () => {
            cancelled = true;
        };
    }, [graphClient]);

    const closePreview = () => {
        setIsPreviewVisible(false);
        setPreviewUrl('');
    };

    const handlePreviewButton = async (fileItem: MicrosoftGraph.DriveItem) => {
        if (!graphClient) return;
        const parentReference = fileItem.parentReference;
        if (!parentReference) {
            return;
        }
        const previewEndpoint = `https://graph.microsoft.com/v1.0/drives/${parentReference.driveId}/items/${fileItem.id}/preview`;
        const response = await graphClient.api(previewEndpoint).post({});
        setPreviewUrl(response.getUrl);
        setIsPreviewVisible(true);
    };

    const FileTemplate = (props) => {
        if (!props.dataContext || !props.dataContext.file) return null;
        const file = props.dataContext.file as MicrosoftGraph.DriveItem;
        return (
            <CustomFile
                fileDetails={file}
                onPreview={() => handlePreviewButton(file)}
            />
        );
    };

    React.useEffect(() => {
        if (fileListRef.current) {
            (fileListRef.current as HTMLElement).addEventListener('itemClick', (e) => {
                const event = e as CustomEvent;
                console.log('FileList item clicked:', event.detail);
                const file = document.querySelector('mgt-file') as any;
                if (file) {
                    file.fileDetails = event.detail;
                }
            });
        }
    }, [selectedChannelName]);

    return (
        <>
            <PageHeader
                title='Channel Files'
                description='View files from access channels you are a member of'
            ></PageHeader>
            <div className={styles.container}>
                <div className={styles.teamChannel}>
                    {loading ? <Loading /> :
                        <Tree aria-label='TeamTree'>
                            {teams.map((team) => (
                                <TreeItem itemType='branch' key={team.id}>
                                    <TreeItemLayout>
                                        <div style={{ display: 'flex' }}>
                                            <span style={{ display: 'flex', alignItems: 'center' }}>
                                                <TeamImg teamId={team.id} graphClient={graphClient} />
                                            </span>
                                            <span style={{ display: 'flex', alignItems: 'center' }}>
                                                {team.displayName}
                                            </span>
                                        </div>
                                    </TreeItemLayout>
                                    <ChannelsTree
                                        teamId={team.id}
                                        graphClient={graphClient}
                                        selectedTeamId={selectedTeamId}
                                        setSelectedTeamId={setSelectedTeamId}
                                        setSelectedChannelName={setSelectedChannelName}
                                    />
                                </TreeItem>
                            ))}
                        </Tree>
                    }
                </div>
                <div className={styles.divider}></div>

                  {selectedChannelName !== '' ? (
                    <FileList
                        ref={fileListRef}
                        key={`${selectedTeamId}-${selectedChannelName}`}
                        accessToken={accessToken}
                        groupId={selectedTeamId}
                        itemPath={selectedChannelName}
                        pageSize={100}
                        className={styles.channelFiles}
                        enableFileUpload={true}
                        disableOpenOnClick={true}
                        onLoadingChange={setFileListLoading} // 需要 FileList 支持此回调
                    >
                        {fileListLoading && <Loading template='loading'></Loading>}
                        <FileTemplate template='file'></FileTemplate>
                    </FileList>
                ) : null}
            </div>
            {isPreviewVisible && (
                <div style={{
                    position: 'fixed',
                    top: 0,
                    left: 0,
                    width: '100%',
                    height: '100%',
                    backgroundColor: 'rgba(0, 0, 0, 0.5)',
                    zIndex: 1000
                }}>
                    <Button
                        appearance="primary"
                        onClick={closePreview}
                        style={{
                            position: 'absolute',
                            top: '5px',
                            right: '5px',
                            zIndex: 1001
                        }}
                    >
                        Close
                    </Button>
                    <div style={{
                        position: 'absolute',
                        top: '50px',
                        left: '50px',
                        right: '50px',
                        bottom: '50px',
                        backgroundColor: 'white',
                        overflow: 'auto'
                    }}>
                        <iframe title="File preview" src={previewUrl} style={{ width: '100%', height: '100%', border: 'none' }}></iframe>
                    </div>
                </div>
            )}
        </>
    );
}

//utils
const blobToBase64 = (blob: Blob) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = reject;
        reader.onload = () => {
            resolve(reader.result as string);
        };
        reader.readAsDataURL(blob);
    });
};

const getAllMyTeams = async (graph: Client) => {
    const teams = await graph
        .api('/me/joinedTeams')
        .select(['displayName', 'id'])
        .get();
    let apiCon = [{
        api: "https://graph.microsoft.com/beta//me/joinedTeams$select=displayName,id/",
        type: "GET"
    }];
    PubSub.publish("Calendar", apiCon);
    return teams?.value || [];
};

const getTeamPhoto = async (graph: Client, teamId: string) => {
    const response = (await graph
        .api(`/teams/${teamId}/photo/$value`)
        .responseType(ResponseType.RAW)
        .get()
    ) as Response;
    let apiCon = [{
        api: "https://graph.microsoft.com/beta/teams/" + teamId + "/photo/$value/",
        type: "GET"
    }];
    PubSub.publish("Calendar", apiCon);
    const blob = await blobToBase64(await response.blob());
    return blob;
};

const getChannelsByTeam = async (graph: Client, teamId: string) => {
    const channels = await graph
        .api(`/teams/${teamId}/channels`)
        .get();
    let apiCon = [{
        api: "https://graph.microsoft.com/beta//teams/" + teamId + "/channels",
        type: "GET"
    }];
    PubSub.publish("Calendar", apiCon);
    return channels?.value || [];
};