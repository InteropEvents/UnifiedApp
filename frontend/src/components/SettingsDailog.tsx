import React, { useState, useEffect } from "react";
import { Dialog, DialogSurface, DialogBody, DialogTitle, DialogTrigger, Tab, TabList, Checkbox } from "@fluentui/react-components";
import { ThemeSwitcher } from './ThemeSwitcher';
import { IconButton } from '@fluentui/react/lib/Button';
import { SettingsDialogProps } from './Interface';
import { useSettingsDialogStyles } from '../styles/Styles';

export const SettingsDialog: React.FC<SettingsDialogProps> = ({ setHandleRemoveAPI }) => {
  const [selectedTab, setSelectedTab] = useState("enableLogFlyout");
  const [isLogFlyoutEnabled, setIsLogFlyoutEnabled] = useState(false);
  const [openNatively, setOpenNatively] = useState(false);
  const [openInBrowser, setOpenInBrowser] = useState(true);
  const classes = useSettingsDialogStyles();

  const handleCheckboxChange = () => {
    setIsLogFlyoutEnabled(!isLogFlyoutEnabled);
  };

  const handleOpenNativelyChange = () => {
    setOpenNatively(true);
    setOpenInBrowser(false);
  };

  const handleOpenInBrowserChange = () => {
    setOpenNatively(false);
    setOpenInBrowser(true);
  };

  useEffect(() => {
    setHandleRemoveAPI(isLogFlyoutEnabled);
  }, [isLogFlyoutEnabled, setHandleRemoveAPI]);

  return (
    <Dialog>
      <DialogTrigger disableButtonEnhancement>
        <IconButton
          iconProps={{ iconName: 'Settings' }}
          title="Settings"
          ariaLabel="Settings"
          className={classes.button}
        />
      </DialogTrigger>

      <DialogSurface>
        <DialogBody>
          <DialogTitle>Settings</DialogTitle>
          <div className={classes.settingsContainer}>
            <div className={classes.sidebar}>
              <TabList selectedValue={selectedTab} onTabSelect={(_, data) => setSelectedTab(data.value as string)} vertical>
                <Tab value="enableLogFlyout">Show logging flyout</Tab>
                <Tab value="theme">Change Theme</Tab>
                <Tab value="files">Files & links</Tab>
              </TabList>
            </div>
            <div className={classes.content}>
              {selectedTab === "enableLogFlyout" && (
                <div>
                  <Checkbox
                    checked={isLogFlyoutEnabled}
                    onChange={handleCheckboxChange}
                    className={classes.roundCheckbox}
                  />
                </div>
              )}
              {selectedTab === "theme" && (
                <div>
                  <ThemeSwitcher />
                </div>
              )}
              {selectedTab === "files" && (
                <div className={classes.checkboxContainer}>
                  <Checkbox
                    label="Open Natively"
                    checked={openNatively}
                    onChange={handleOpenNativelyChange}
                    className={classes.roundCheckbox}
                  />
                  <Checkbox
                    label="Open in browser"
                    checked={openInBrowser}
                    onChange={handleOpenInBrowserChange}
                    className={classes.roundCheckbox}
                  />
                </div>
              )}
            </div>
          </div>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};