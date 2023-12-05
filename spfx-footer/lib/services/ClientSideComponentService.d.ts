import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IMessageBannerProperties } from "../models/IMessageBannerProperties";
import { IHostProperties } from "../models/IHostProperties";
declare class ClientSideComponentService {
    private _context;
    constructor(context: ApplicationCustomizerContext);
    setProperties: (properties?: IMessageBannerProperties, hostProperties?: IHostProperties) => Promise<void>;
    private _getCustomActionByComponentId;
}
export default ClientSideComponentService;
//# sourceMappingURL=ClientSideComponentService.d.ts.map