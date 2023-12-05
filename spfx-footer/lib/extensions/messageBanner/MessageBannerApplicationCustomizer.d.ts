import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { IMessageBannerProperties } from '../../models/IMessageBannerProperties';
/** A Custom Action which can be run during execution of a Client Side Application */
export default class MessageBannerApplicationCustomizer extends BaseApplicationCustomizer<IMessageBannerProperties> {
    private _topPlaceholder;
    private _extensionProperties;
    private _clientSideComponentService;
    onInit(): Promise<void>;
    /**
     * Event handler that fires on every page load
     */
    private onNavigated;
    /**
     * Render the 'content viewable by external users' banner on the current page
     */
    private renderBanner;
    onDispose(): void;
}
//# sourceMappingURL=MessageBannerApplicationCustomizer.d.ts.map