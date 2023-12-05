var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
var useState = React.useState, useEffect = React.useEffect;
import styles from './Banner.module.scss';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import BannerPanel from '../BannerPanel/BannerPanel';
import * as strings from 'MessageBannerApplicationCustomizerStrings';
import { SPPermission } from '@microsoft/sp-page-context';
import isPast from 'date-fns/isPast';
import formatDate from 'date-fns/format';
import { Text } from '@microsoft/sp-core-library';
import { DEFAULT_PROPERTIES } from '../../../../models/IMessageBannerProperties';
var BANNER_CONTAINER_ID = 'CustomMessageBannerContainer';
var Banner = function (props) {
    var _a = useState(props.settings), defaultSettings = _a[0], setDefaultSettings = _a[1];
    var _b = useState(props.settings), settings = _b[0], setSettings = _b[1];
    var _c = useState(false), isPanelOpen = _c[0], setIsPanelOpen = _c[1];
    var _d = useState(false), isSaving = _d[0], setIsSaving = _d[1];
    useEffect(function () {
        // Adjust pre allocated parent container height for previewing
        if (props.settings.enableSetPreAllocatedTopHeight) {
            document.getElementById(BANNER_CONTAINER_ID).parentElement.style.height = "".concat(settings.bannerHeightPx, "px");
        }
    }, [settings.bannerHeightPx]);
    var visibleStartDate = settings.visibleStartDate ? new Date(settings.visibleStartDate) : null;
    var isPastVisibleStartDate = settings.visibleStartDate && isPast(visibleStartDate);
    var isCurrentUserAdmin = props.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
    // Set Panel to open
    var handleOpenClick = function () {
        setIsPanelOpen(true);
    };
    // handle cancel button, and discard all the input
    var handleCancelOrDismiss = function () {
        if (!isSaving) {
            setIsPanelOpen(false);
            setSettings(defaultSettings); //return to original settings
        }
    };
    // handle save button, and update the banner based on the input
    var handleSave = function () { return __awaiter(void 0, void 0, void 0, function () {
        var hostProperties, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    setIsSaving(true);
                    hostProperties = {};
                    // Set host property 'preAllocatedApplicationCustomizerTopHeight' when saving custom action properties
                    if (props.settings.enableSetPreAllocatedTopHeight) {
                        hostProperties.preAllocatedApplicationCustomizerTopHeight = settings.bannerHeightPx;
                    }
                    return [4 /*yield*/, props.clientSideComponentService.setProperties(settings, hostProperties)];
                case 1:
                    _a.sent();
                    setIsPanelOpen(false);
                    setIsSaving(false);
                    setDefaultSettings(settings);
                    return [3 /*break*/, 3];
                case 2:
                    error_1 = _a.sent();
                    console.log("Unable to set custom action properties. ".concat(error_1.message), error_1);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    // handle the changes
    var handleFieldChange = function (newSetting) {
        var newSettings = __assign(__assign({}, settings), newSetting);
        setSettings(newSettings);
    };
    // handle the reset to default link, to revert the changes
    var resetToDefaults = function () {
        var mergedDefaultSettings = __assign(__assign({}, settings), DEFAULT_PROPERTIES);
        setSettings(mergedDefaultSettings);
    };
    // handle url is user input url in the text area
    var parseTokens = function (textWithTokens, context) {
        var tokens = [
            { token: '{siteUrl}', value: context.pageContext.site.absoluteUrl },
            { token: '{webUrl}', value: context.pageContext.web.absoluteUrl },
        ];
        var outputText = tokens.reduce(function (text, tokenItem) {
            return text.replace(tokenItem.token, tokenItem.value);
        }, textWithTokens);
        return outputText;
    };
    //If there is a future start date and it hasn't yet occurred,
    // and either the current user isn't an admin or the user is an admin but the disableSiteAdminUI flag is set,
    // then render nothing
    if (visibleStartDate && !isPastVisibleStartDate && (!isCurrentUserAdmin || settings.disableSiteAdminUI))
        return null;
    return (React.createElement("div", { id: BANNER_CONTAINER_ID },
        React.createElement("div", { className: styles.BannerContainer, style: { height: settings.bannerHeightPx } },
            !settings.disableSiteAdminUI && isCurrentUserAdmin && !!visibleStartDate && (isPastVisibleStartDate
                ? React.createElement("div", { className: styles.AdminUserVisibilityBadge }, strings.BannerBadgeIsVisibleToUsersMessage)
                : React.createElement("div", { className: styles.AdminUserVisibilityBadge }, Text.format(strings.BannerBadgeNotVisibleToUsersMessage, formatDate(visibleStartDate, 'PPPP')))),
            React.createElement("div", { dangerouslySetInnerHTML: { __html: parseTokens(settings.message, props.context) }, style: { color: settings.textColor, fontSize: settings.textFontSizePx } }),
            !settings.disableSiteAdminUI && isCurrentUserAdmin && (React.createElement(IconButton, { iconProps: { iconName: 'Edit', styles: { root: { color: settings.textColor } } }, onClick: handleOpenClick })),
            !settings.disableSiteAdminUI && (React.createElement(BannerPanel, { isOpen: isPanelOpen, isSaving: isSaving, onCancelOrDismiss: handleCancelOrDismiss, onFieldChange: handleFieldChange, onSave: handleSave, resetToDefaults: resetToDefaults, settings: settings })))));
};
export default Banner;
//# sourceMappingURL=Banner.js.map