import { makeStyles, shorthands, tokens } from '@fluentui/react-components';

export const useTeamsPageStyles = makeStyles({
    container: {
        display: 'flex',
        flexDirection: 'row',
    },

    teamChannel: {
        display: 'flex',
        flexDirection: 'column',
        flexWrap: 'nowrap',
        width: '30%'
    },

    divider: {
        width: '1px',
        backgroundColor: '#D9D9D9'
    },

    channelFiles: {
        display: 'flex',
        flexDirection: 'column',
        flexWrap: 'nowrap',
        width: '69%',
        '--file-list-box-shadow': 'none'
    }
});

export const usePresenceGetStyles = makeStyles({
    simpleLogin: {
    '--person-avatar-size': '32px'
    },
    available: {
    marginRight: '5px',
    color: 'green',
    },
    busy: {
    marginRight: '5px',
    color: 'red',
    },
    away: {
    marginRight: '5px',
    color: '#FFA500',
    },
    offline: {
    marginRight: '5px',
    color: 'gray',
    },
    presenceContainer: {
        display: 'flex',
        alignItems: 'center',
        marginRight: '15px',
        fontSize: '13px',
        marginLeft: '58px',
        width: '100%',
        justifyContent: "space-between",
    },
    presenceInfo:{
        display: 'flex',
        alignItems: 'center',
    }, 
    popupContent: {
        padding: '10px',
        display: 'flex',
        flexDirection: 'column',
        backgroundColor: 'white', 
        position: 'fixed',
        border: '1px solid #ccc',
        boxShadow: '0 4px 8px rgba(15, 8, 8, 0.1)', 
        borderRadius: '8px', 
        width: '150px',
    },
    menuList: {
        fontSize: '14px', 
        fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, "Apple Color Emoji", "Segoe UI Emoji", sans-serif',
    },
    menuButton: {
        background: "none",
        border: "none",
        padding: "0",
        margin: "0",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
      },
});

export const useFlyoutLoginStyles = makeStyles({
  personAvatar: {
    '--person-avatar-size': '48px'
  },
  container: {
    display: 'flex',
    alignItems: 'flex-start',
    flexDirection : 'column',
  },
  statusContainer:{
    display: 'flex',
    width: '100%',
    marginTop: '10px'
  }
});
export const useEntraIDStyles = makeStyles({
    container: {
      height: '80vh',
      display: 'flex',
      flexDirection: 'column',
    },
    nav: {
      width: '150px',
      borderRight: '1px solid #ccc',
      padding: '1rem',
    },
    navItem: {
      padding: '0.5rem 1rem',
      cursor: 'pointer',
      background: 'transparent',
      color: '#000',
      '&.active': {
        background: '#0078d4',
        color: '#fff',
      },
    },
    content: {
      flex: 1,
      padding: '1rem',
    },
    main: {
      flex: 1,
      display: 'flex',
    },
  });
  
  export const useUserStyles = makeStyles({
  dialogBody: {
      display: 'flex',
      flexDirection: 'column', 
      maxHeight: '80vh', 
      overflowY: 'auto',
  
      },
  subcontainer: {
        display: 'flex',
        flexDirection: 'column',
        maxWidth: '700px',
        },
    listContainer: {
        listStyle: 'none',
        padding: 0,
        margin: 0,
      },
    listItem: {
      display: 'flex',
      alignItems: 'center',
      marginBottom: '10px',
    },
    avatar: {
      width: '40px',
      height: '40px',
      borderRadius: '50%',
      marginRight: '10px',
    },
    placeholderAvatar: {
      width: '40px',
      height: '40px',
      borderRadius: '50%',
      backgroundColor: '#ccc',
      color: '#fff',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontWeight: 'bold',
      marginRight: '10px',
    },
    userLink: {
      color: 'blue',
      textDecoration: 'underline',
      cursor: 'pointer',
    },
    userListContainer: {
      maxHeight: '80vh', 
      overflowY: 'auto',
    },
    required: {
      color: 'red',
      marginLeft: '4px',
    },
    filterInput: {
      marginBottom: '1rem',
      width: '50%',
    },
    twoColumnRow: {
      display: 'grid',
      gridTemplateColumns: '1fr 2fr', // controls the width of the columns
      alignItems: 'center',
      gap: '1rem', 
      marginBottom: '1rem', 
    },
    label: {
      textAlign: 'left', 
      paddingRight: '1rem', 
    },
    tabButton: {
      marginRight: '1rem',
      backgroundColor: 'transparent',
      border: 'none',
      borderBottom: '2px solid transparent',
      padding: '0.5rem 1rem',
      cursor: 'pointer',
      '&:hover': {
        borderBottom: '2px solid lightblue',
      },
      '&:focus': {
        borderBottom: '4px solid blue',
      },
    },
    button: {
      height: 'auto',
      padding: '1rem 1rem',
      fontSize: 'inherit',
    },
    table: {
      width: '90%',
      borderCollapse: 'collapse',
      marginTop: '1rem',
      tableLayout: 'auto',
    },
    th: {
      textAlign: 'left',
      borderBottom: '1px solid #ccc',
      padding: '0.5rem',
    },
    td: {
      padding: '0.5rem',
      borderBottom: '1px solid #eee',
    },
});

export const useDirectReportsStyles = makeStyles({
  toolbar: {
    justifyContent: 'space-between'
  }
});

export const useOutlookPageStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'row'
  },
  panels: {
    ...shorthands.padding('10px')
  },
  main: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    flexGrow: 1,
    ...shorthands.padding('0', '15px')
  },
  // Add vertical tablist styles
  verticalTabList: {
    display: 'flex',
    flexDirection: 'column',
    width: '200px',
    borderRight: '1px solid #e0e0e0',
    ...shorthands.padding('10px', '0'),
    height: '100%'
  },
  // Style for individual tabs
  verticalTab: {
    justifyContent: 'flex-start',
    textAlign: 'left',
    ...shorthands.padding('8px', '16px'),
    ...shorthands.margin('2px', '0')
  },
  // Modified container for vertical layout
  verticalContainer: {
    display: 'flex',
    flexDirection: 'row',
    height: 'calc(100vh - 120px)' // Adjust based on your header height
  },
  // Content area - right side
  contentArea: {
    flexGrow: 1,
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.padding('10px')
  },
  // Message list panel - moved to middle
  messageList: {
    display: 'flex',
    flexDirection: 'column',
    width: '300px',
    minWidth: '300px',
    borderRight: '1px solid #e0e0e0',
    ...shorthands.padding('10px', '0'),
    height: '100%',
    overflowY: 'auto'
  },
  emptyState: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
    fontSize: '16px',
    color: '#666666',
    fontStyle: 'italic'
  },
  messageItem: {
    cursor: 'pointer',
    transition: 'background-color 0.1s',
    ':hover': {
      backgroundColor: 'var(--button-background--hover)'
    }
  },
  buttonContainer: {
    display: 'flex',
    gap: '8px',
  },
  newMailButton: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    boxShadow: '0 2px 4px rgba(0, 0, 0, 0.1)',
    ':hover': {
      backgroundColor: tokens.colorBrandBackgroundHover,
    }
  },
  readUnreadButton: {
    backgroundColor: 'transparent',
    ':hover': {
      boxShadow: '0 2px 4px rgba(0, 0, 0, 0.1)'
    }
  }
});

export const useCalendarPageStyles = makeStyles({
  container: {
    position: 'relative',
    display: 'flex',
    flexDirection: 'column'
  },
  panels: {
    ...shorthands.padding('10px')
  },
  main: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    width: '100%'
  },
  side: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    width: '100%'
  },
  navigation: {
    display: 'flex',
    flexDirection: 'column'
  },
  mainButton: {
    display: 'flex',
    justifyContent: 'space-between'
  },
  loading: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '4vh'
  },
  loadingMessage: {
    paddingLeft: '10px'
  }
});

export const usePageHeaderStyles = makeStyles({
  divider: {
    alignItems: 'self-start',
    paddingTop: '20px',
    marginBottom: '20px'
  }
});

export const useMailContentStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: '16px'
  },

  headerCard: {
    boxShadow: 'var(--box-shadow)',
    ...shorthands.padding('16px'),
    ...shorthands.margin('8px', '0'),
    borderRadius: '4px',
    backgroundColor: 'var(--background-color)'
  },

  subject: {
    fontSize: '20px',
    fontWeight: '600',
    color: 'var(--color-sub1)',
    ...shorthands.margin('0', '0', '16px', '0')
  },

  senderSection: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.margin('12px', '0'),
    '& mgt-person': {
      '--font-size': '14px',
      '--avatar-size-s': '32px'
    }
  },

  recipientsSection: {
    display: 'flex',
    flexDirection: 'row',  // Change to row for horizontal layout
    alignItems: 'center',  // Align items vertically in the center
    ...shorthands.margin('12px', '0')
  },

  recipientsList: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '8px',
    '& mgt-person': {
      '--font-size': '12px',
      '--avatar-size-s': '24px'
    }
  },

  label: {
    fontSize: '14px',
    fontWeight: '600',
    color: 'var(--color-sub1)',
    width: '80px'
  },

  dateTime: {
    fontSize: '14px',
    color: 'var(--color-sub2)'
  },

  dateTimeContainer: {
    display: 'flex',
    justifyContent: 'flex-end',
    ...shorthands.margin('12px', '0', '0', '0')
  },

  contentCard: {
    boxShadow: 'var(--box-shadow)',
    ...shorthands.padding('16px'),
    ...shorthands.margin('8px', '0'),
    borderRadius: '4px',
    backgroundColor: 'var(--background-color)'
  },

  emailContent: {
    width: '100%',
    fontSize: '14px',
    lineHeight: '1.5',
    color: 'var(--color-sub1)',
    '& img': {
      maxWidth: '100%',
      height: 'auto'
    },
    '& a': {
      color: 'var(--accent-color)'
    }
  },

  emptyContent: {
    fontStyle: 'italic',
    color: 'var(--color-sub2)'
  }
});

export const useHeaderStyles = makeStyles({
  root: {
    '--login-signed-in-hover-background': 'transparent'
  },

  header: {
    display: 'flex',
    flexDirection: 'row',
    flexWrap: 'nowrap',
    width: '100%',
    height: '50px',
    boxSizing: 'border-box',
    alignItems: 'center',
    backgroundColor: tokens.colorNeutralForeground1Static,
    zIndex: '1',
    minHeight: '50px'
  },

  name: {
    color: tokens.colorNeutralForegroundOnBrand,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    paddingLeft: '12px'
  },

  waffleIcon: {
    fontSize: tokens.fontSizeBase600,
    paddingTop: '5px',
    color: tokens.colorNeutralForegroundOnBrand
  },

  waffle: {
    display: 'flex',
    width: 'auto',
    height: '100%',
    boxSizing: 'border-box',
    flexGrow: '1',
    alignItems: 'center',
    minWidth: 0,
    paddingLeft: '8px'
  },

  waffleLogo: {
    ...shorthands.flex('none')
  },

  waffleTitle: {
    ...shorthands.flex('auto'),
    display: 'flex',
    alignItems: 'center',
    minWidth: 0,
    overflow: 'hidden'
  },

  search: {
    display: 'flex',
    width: '100%',
    height: 'auto',
    boxSizing: 'border-box',
    flexGrow: '1',
    justifyContent: 'center'
  },

  searchBox: {
    minWidth: '320px',
    maxWidth: '468px',
    paddingRight: '1em',
    paddingLeft: '1em',
    width: '100%'
  },

  infoIcon: {
    color: tokens.colorNeutralForegroundOnBrand,
    ':hover': {
      color: tokens.colorNeutralForegroundOnBrand,
      ':active': {
        color: tokens.colorNeutralForegroundOnBrand
      }
    }
  },

  login: {
    display: 'flex',
    width: 'auto',
    height: '100%',
    boxSizing: 'border-box',
    flexShrink: 1,
    alignItems: 'center',
    justifyContent: 'flex-end',
    gap: '4px',
    minWidth: 0,
    paddingRight: '4px'
  },

  signedOut: {
    ...shorthands.padding('4px', '8px')
  },

  signedIn: {
    minWidth: 0,
    ...shorthands.padding('0px', '4px')
  }
});

export const useLoadingStyles = makeStyles({
  root: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: 'calc(100vh - 300px)'
  },
  message: {
    paddingLeft: '10px'
  }
});

export const useMessagesStyles = makeStyles({
  email: {
    boxShadow: 'var(--box-shadow)',
    ...shorthands.padding('10px'),
    ...shorthands.margin('8px'),
    ':hover': {
      borderLeftWidth: '4px',
      borderLeftColor: 'var(--input-border-color--hover)',
      borderLeftStyle: 'solid',
      paddingLeft: '6px',
      backgroundColor: '#f5f5f5'
    },
    '& mgt-person': {
      '--font-size': '12px',
      '--avatar-size-s': '16px'
    }
  },

  seletedEmailRadio: {
    // visually hidden but accessible
    position: 'absolute',
    left: '-10000px',
    width: '1px',
    height: '1px',
    ':checked + div, :focus + div': {
      borderLeftWidth: '4px',
      borderLeftColor: 'var(--input-border-color--hover)',
      borderLeftStyle: 'solid',
      paddingLeft: '6px'
    }
  },

  link: {
    color: 'var(--color-sub1)',
    textDecorationLine: 'none'
  },

  header: {
    display: 'flex',
    justifyContent: 'space-between'
  },

  subject: {
    color: 'var(--color-sub1)',
    fontSize: '14px',
    ...shorthands.margin('0'),
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    maxWidth: '200px' // Adjust based on your layout
  },

  unreadSubject: {
    color: '#0078d4',
    fontWeight: '600'
  },

  title: {
    display: 'flex',
    justifyContent: 'space-between',
    marginBottom: '4px',
    color: 'var(--color-sub1)',
    width: '100%'
  },

  date: {
    fontSize: '12px',
    paddingLeft: '4px',
    whiteSpace: 'nowrap'
  },

  body: {
    fontSize: '13px',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    color: 'var(--color-sub2)',
    lineHeight: '1.4em',
    width: '100%'
  },

  emptyBody: {
    fontStyle: 'italic'
  }
});

export const useSettingsDialogStyles = makeStyles({
  settingsContainer: {
    display: "flex",
    flexDirection: "row",
    gap: "1rem",
    width: "100%",
  },
  sidebar: {
    display: "flex",
    flexDirection: "column",
    minWidth: "200px",
    padding: "1rem",
    borderRight: "1px solid #ccc",
  },
  content: {
    flexGrow: 1,
    padding: "1rem",
    minWidth: "300px",
  },
  button: {
    width: '40px',
    minWidth: '40px',
    height: '40px',
    margin: 0,
    padding: 0,
    fontSize: '16px',
    color: 'white',
    backgroundColor: 'transparent',
    ':hover': {
      color: 'white',
      backgroundColor: 'rgba(255, 255, 255, 0.1)',
    },
  },
  select: {
    width: "100%",
    padding: "0.75rem",
    fontSize: "16px",
    borderRadius: "6px",
    border: "1px solid #ccc",
  },
  checkboxContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "0.5rem",
  },
  roundCheckbox: {
    '& .ms-Checkbox-checkbox': {
      borderRadius: '50%',
    },
    '& .ms-Checkbox-checkmark': {
      borderRadius: '50%',
    },
  },
});

export const useSideNavigationStyles = makeStyles({
  tab: {
    paddingTop: '12px',
    paddingBottom: '12px'
  },
  activeTab: {
    backgroundColor: tokens.colorSubtleBackgroundHover
  }
});

export const useChannelFilesStyles = makeStyles({
  fileGrid: {
    paddingBottom: '10px'
  }
});

export const useSiteFilesStyles = makeStyles({
  picker: {
    paddingBottom: '10px',
    display: 'block'
  }
});

export const useFilesResultsStyles = makeStyles({
  container: {
    ...shorthands.gap('16px'),
    display: 'flex',
    flexDirection: 'row',
    flexWrap: 'wrap'
  },
  card: {
    width: '300px',
    height: 'fit-content',
    maxWidth: '100%'
  },
  caption: {
    color: tokens.colorNeutralForeground3
  },
  noDataSearchTerm: {
    fontWeight: tokens.fontWeightSemibold
  },
  emptyContainer: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    flexDirection: 'column',
    height: 'calc(100vh - 300px)'
  },
  fileContainer: {
    display: 'flex'
  },
  fileTitle: {
    paddingLeft: '10px',
    alignSelf: 'center'
  },
  noDataMessage: {
    paddingLeft: '10px'
  },
  noDataIcon: {
    fontSize: '128px'
  },
  row: {
    cursor: 'pointer'
  }
});

export const useTaxonomyExplorerStyles = makeStyles({
  main: {
    ...shorthands.gap('36px'),
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'wrap'
  },

  title: {
    ...shorthands.margin(0, 0, '12px')
  },

  description: {
    ...shorthands.margin(0, 0, '12px')
  },

  card: {
    width: '480px',
    maxWidth: '100%',
    height: 'fit-content',
    ...shorthands.margin('12px', 0)
  },

  caption: {
    color: tokens.colorNeutralForeground3
  },

  icon: {
    width: '24px',
    height: '24px'
  },

  text: {
    ...shorthands.margin(0)
  },

  groupPanel: {
    ...shorthands.margin('12px', '24px')
  },

  termPanel: {
    ...shorthands.margin('12px', '36px')
  }
});

export const useDashboardPageStyles = makeStyles({
  panels: {
    ...shorthands.padding('10px')
  }
});

export const useLayoutStyles = makeStyles({
  sidebar: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    height: '100%',
    minWidth: '295px',
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground6
  },
  main: {
    backgroundColor: tokens.colorNeutralBackground1,
    display: 'flex',
    flexDirection: 'row',
    width: 'auto',
    height: 'calc(100vh - 50px)',
    boxSizing: 'border-box'
  },
  minimized: {
    minWidth: 'auto'
  },
  page: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap'
  },
  content: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    width: '100%',
    height: 'auto',
    boxSizing: 'border-box',
    ...shorthands.margin('10px'),
    ...shorthands.overflow('auto')
  }
});

export const useConditionalAccessStyles = makeStyles({
  sectionHeader: {
    borderBottom: '1px solid #ccc',
    paddingBottom: '0.5rem',
  },
  dialogBody: {
    display: 'flex',
    flexDirection: 'column',
    maxHeight: '80vh',
    overflowY: 'auto',
  },
  subcontainer: {
    display: 'flex',
    flexDirection: 'column',
  },
  listContainer: {
    listStyle: 'none',
    padding: 0,
    margin: 0,
  },
  link: {
    color: 'rgb(0, 120, 212)',
    textDecoration: 'underline',
    cursor: 'pointer',
  },
  createdText: {
    marginLeft: '10px',
  },
  required: {
    color: 'red',
    marginLeft: '4px',
  },
  linkButton: {
    color: 'rgb(0, 120, 212)',
    textDecoration: 'underline',
    cursor: 'pointer',
    fontSize: 'inherit',
    display: 'inline',
  },
  button: {
    height: 'auto',
    padding: '1rem 1rem',
    fontSize: 'inherit',
  },
  table: {
    width: '90%',
    borderCollapse: 'collapse',
    marginTop: '1rem',
    tableLayout: 'auto',
  },
  th: {
    textAlign: 'left',
    borderBottom: '1px solid #ccc',
    padding: '0.5rem',
  },
  td: {
    padding: '0.5rem',
    borderBottom: '1px solid #eee',
  },
  deleteButton: {
    backgroundColor: 'transparent',
    border: 'none',
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem',
    cursor: 'pointer',
    fontSize: 'inherit',
    padding: '0',
    color: 'rgb(0, 120, 212)',
    textDecoration: 'none',
  },
  deleteIcon: {
    color: 'inherit',
  },
  submitButton: {
    marginTop: '1rem',
    backgroundColor: 'rgb(0, 120, 212)',
    color: 'white',
    cursor: 'pointer',
    border: 'none',
    padding: '0.5rem 1rem',
    fontSize: '1rem',
    borderRadius: '4px',
    '&:disabled': {
      backgroundColor: '#ccc',
      cursor: 'not-allowed',
    },
  },
});