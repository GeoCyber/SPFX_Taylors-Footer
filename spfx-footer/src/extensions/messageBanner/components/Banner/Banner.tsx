import * as React from 'react';
const { useState, useEffect } = React;
import { IBannerProps } from './IBannerProps';
import styles from './Banner.module.scss';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import BannerPanel from '../BannerPanel/BannerPanel';
import * as strings from 'MessageBannerApplicationCustomizerStrings';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { SPPermission } from '@microsoft/sp-page-context';
import isPast from 'date-fns/isPast';
import formatDate from 'date-fns/format';
import { Text } from '@microsoft/sp-core-library';

import { DEFAULT_PROPERTIES } from '../../../../models/IMessageBannerProperties';
import { IHostProperties } from '../../../../models/IHostProperties';


const BANNER_CONTAINER_ID = 'CustomMessageBannerContainer';

const Banner = (props: IBannerProps) => {
  const [defaultSettings, setDefaultSettings] = useState(props.settings);
  const [settings, setSettings] = useState(props.settings);
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    // Adjust pre allocated parent container height for previewing
    if (props.settings.enableSetPreAllocatedTopHeight) {
      document.getElementById(BANNER_CONTAINER_ID).parentElement.style.height = `${settings.bannerHeightPx}px`;
    }
  }, [settings.bannerHeightPx]);

  const visibleStartDate = settings.visibleStartDate ? new Date(settings.visibleStartDate) : null;
  const isPastVisibleStartDate = settings.visibleStartDate && isPast(visibleStartDate);
  const isCurrentUserAdmin = props.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);

  // Set Panel to open
  const handleOpenClick = (): void => {
    setIsPanelOpen(true);
  };

  // handle cancel button, and discard all the input
  const handleCancelOrDismiss = (): void => {
    if (!isSaving) {
      setIsPanelOpen(false);
      setSettings(defaultSettings); //return to original settings
    }
  };

  // handle save button, and update the banner based on the input
  const handleSave = async (): Promise<void> => {
    try {
      setIsSaving(true);
      const hostProperties : IHostProperties = {};
      // Set host property 'preAllocatedApplicationCustomizerTopHeight' when saving custom action properties
      if (props.settings.enableSetPreAllocatedTopHeight) {
          hostProperties.preAllocatedApplicationCustomizerTopHeight = settings.bannerHeightPx;
      }
      await props.clientSideComponentService.setProperties(settings, hostProperties);
      setIsPanelOpen(false);
      setIsSaving(false);
      setDefaultSettings(settings);
    }
    catch (error) {
      console.log(`Unable to set custom action properties. ${error.message}`, error);
    }
  };

  // handle the changes
  const handleFieldChange = (newSetting: {[ key: string ]: unknown }): void => {
    const newSettings = { ...settings, ...newSetting };
    setSettings(newSettings);
  };

  // handle the reset to default link, to revert the changes
  const resetToDefaults = (): void => {
    const mergedDefaultSettings = { ...settings, ...DEFAULT_PROPERTIES };
    setSettings(mergedDefaultSettings);
  };

  // handle url is user input url in the text area
  const parseTokens = (textWithTokens: string, context: BaseComponentContext): string => {
    const tokens = [
      { token: '{siteUrl}', value: context.pageContext.site.absoluteUrl },
      { token: '{webUrl}', value: context.pageContext.web.absoluteUrl },
    ];

    const outputText = tokens.reduce((text, tokenItem) => {
      return text.replace(tokenItem.token, tokenItem.value);
    }, textWithTokens);

    return outputText;
  };


  //If there is a future start date and it hasn't yet occurred,
  // and either the current user isn't an admin or the user is an admin but the disableSiteAdminUI flag is set,
  // then render nothing
  if (visibleStartDate && !isPastVisibleStartDate && (!isCurrentUserAdmin || settings.disableSiteAdminUI)) return null;

  return (
    <div id={BANNER_CONTAINER_ID}>
      {/* Banner*/}
      <div className={styles.BannerContainer} style={{ height: settings.bannerHeightPx }}>
        {!settings.disableSiteAdminUI && isCurrentUserAdmin && !!visibleStartDate && (isPastVisibleStartDate
          ? <div className={styles.AdminUserVisibilityBadge}>{strings.BannerBadgeIsVisibleToUsersMessage}</div>
          : <div className={styles.AdminUserVisibilityBadge}>{Text.format(strings.BannerBadgeNotVisibleToUsersMessage, formatDate(visibleStartDate, 'PPPP'))}</div>
        )}
        <div
          dangerouslySetInnerHTML={{__html: parseTokens(settings.message, props.context)}}
          style={{ color: settings.textColor, fontSize: settings.textFontSizePx }} />
        {/* Edit button, display based on user role*/}
        {!settings.disableSiteAdminUI && isCurrentUserAdmin && (
          <IconButton
            iconProps={{ iconName: 'Edit', styles: { root: { color: settings.textColor}}}}
            onClick={handleOpenClick}
          />
        )}
        {/* Panel*/}
        {!settings.disableSiteAdminUI && (<BannerPanel
            isOpen={isPanelOpen}
            isSaving={isSaving}
            onCancelOrDismiss={handleCancelOrDismiss}
            onFieldChange={handleFieldChange}
            onSave={handleSave}
            resetToDefaults={resetToDefaults}
            settings={settings}
          />
        )}
      </div>
    </div>
  );
};

export default Banner;
