var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
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
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
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
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import * as strings from 'MessageBannerApplicationCustomizerStrings';
import Banner from './components/Banner/Banner';
import ClientSideComponentService from '../../services/ClientSideComponentService';
import { DEFAULT_PROPERTIES } from '../../models/IMessageBannerProperties';
var LOG_SOURCE = 'MessageBannerApplicationCustomizer';
/** A Custom Action which can be run during execution of a Client Side Application */
var MessageBannerApplicationCustomizer = /** @class */ (function (_super) {
    __extends(MessageBannerApplicationCustomizer, _super);
    function MessageBannerApplicationCustomizer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    MessageBannerApplicationCustomizer.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                Log.info(LOG_SOURCE, "Initialized ".concat(strings.Title));
                //const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
                // Init services
                this._clientSideComponentService = new ClientSideComponentService(this.context);
                // Merge passed properties with default properties, overriding any defaults
                this._extensionProperties = __assign(__assign({}, DEFAULT_PROPERTIES), this.properties);
                // Don't show banner if message is empty
                if (!this._extensionProperties.message) {
                    Log.info(LOG_SOURCE, "Skip rendering. No banner message configured.");
                    return [2 /*return*/];
                }
                //Event handler to re-render banner on each page navigation
                this.context.application.navigatedEvent.add(this, this.onNavigated);
                return [2 /*return*/];
            });
        });
    };
    /**
     * Event handler that fires on every page load
     */
    MessageBannerApplicationCustomizer.prototype.onNavigated = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                this.renderBanner();
                return [2 /*return*/];
            });
        });
    };
    /**
     * Render the 'content viewable by external users' banner on the current page
     */
    MessageBannerApplicationCustomizer.prototype.renderBanner = function () {
        if (!this._topPlaceholder) {
            this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
            if (!this._topPlaceholder) {
                Log.error(LOG_SOURCE, new Error("Unable to render Top placeholder"));
                return;
            }
        }
        //Render Banner React component
        var bannerProps = {
            context: this.context,
            settings: this._extensionProperties,
            clientSideComponentService: this._clientSideComponentService
        };
        var bannerComponent = React.createElement(Banner, bannerProps);
        //ReactDom.render(bannerComponent, document.getElementById('CommentsWrapper')); // replace commentsWrapper with footer
        ReactDom.render(bannerComponent, document.getElementById('CommentsWrapper')); // replace commentsWrapper with footer
    };
    MessageBannerApplicationCustomizer.prototype.onDispose = function () {
        if (this._topPlaceholder) {
            this._topPlaceholder.dispose();
        }
    };
    __decorate([
        override
    ], MessageBannerApplicationCustomizer.prototype, "onInit", null);
    __decorate([
        override
    ], MessageBannerApplicationCustomizer.prototype, "onDispose", null);
    return MessageBannerApplicationCustomizer;
}(BaseApplicationCustomizer));
export default MessageBannerApplicationCustomizer;
//# sourceMappingURL=MessageBannerApplicationCustomizer.js.map