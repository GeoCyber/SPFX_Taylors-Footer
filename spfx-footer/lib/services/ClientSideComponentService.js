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
import { SPHttpClient } from "@microsoft/sp-http";
var ClientSideComponentService = /** @class */ (function () {
    function ClientSideComponentService(context) {
        var _this = this;
        this.setProperties = function (properties, hostProperties) { return __awaiter(_this, void 0, void 0, function () {
            var componentId, customAction, body, error_1, errorMessage;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        componentId = this._context.manifest.id;
                        return [4 /*yield*/, this._getCustomActionByComponentId(componentId)];
                    case 1:
                        customAction = _a.sent();
                        if (!customAction)
                            return [2 /*return*/];
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 4, , 5]);
                        body = {};
                        if (properties)
                            body["ClientSideComponentProperties"] = JSON.stringify(properties);
                        if (hostProperties)
                            body["HostProperties"] = JSON.stringify(hostProperties);
                        return [4 /*yield*/, this._context.spHttpClient.post(customAction["@odata.id"], SPHttpClient.configurations.v1, {
                                headers: {
                                    "X-HTTP-Method": "MERGE",
                                    "content-type": "application/json; odata=nometadata"
                                },
                                body: JSON.stringify(body)
                            })];
                    case 3:
                        _a.sent();
                        return [3 /*break*/, 5];
                    case 4:
                        error_1 = _a.sent();
                        errorMessage = "Unable to update custom action with componentId ".concat(componentId, ". ").concat(error_1.message);
                        console.log("ERROR: ".concat(errorMessage));
                        throw new Error(errorMessage);
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        this._getCustomActionByComponentId = function (componentId) { return __awaiter(_this, void 0, void 0, function () {
            var result, customActionFilter, webCustomActionUrl, siteCustomActionUrl, _a, webScopeResponse, siteScopeResponse, webResult, siteResult, error_2;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        result = null;
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 10, , 11]);
                        customActionFilter = "$filter=ClientSideComponentId eq guid'".concat(componentId, "'");
                        webCustomActionUrl = "".concat(this._context.pageContext.web.absoluteUrl, "/_api/Web/UserCustomActions?").concat(customActionFilter);
                        siteCustomActionUrl = "".concat(this._context.pageContext.site.absoluteUrl, "/_api/Site/UserCustomActions?").concat(customActionFilter);
                        return [4 /*yield*/, Promise.all([
                                this._context.spHttpClient.get(webCustomActionUrl, SPHttpClient.configurations.v1),
                                this._context.spHttpClient.get(siteCustomActionUrl, SPHttpClient.configurations.v1)
                            ])];
                    case 2:
                        _a = _b.sent(), webScopeResponse = _a[0], siteScopeResponse = _a[1];
                        if (!webScopeResponse.ok) return [3 /*break*/, 4];
                        return [4 /*yield*/, webScopeResponse.json()];
                    case 3:
                        webResult = _b.sent();
                        result = webResult && webResult.value.length > 0 ? webResult.value[0] : null;
                        return [3 /*break*/, 5];
                    case 4: throw new Error("Unable to check web-scoped custom actions. ".concat(webScopeResponse.status, " ").concat(webScopeResponse.statusText));
                    case 5:
                        if (!siteScopeResponse.ok) return [3 /*break*/, 8];
                        if (!!result) return [3 /*break*/, 7];
                        return [4 /*yield*/, siteScopeResponse.json()];
                    case 6:
                        siteResult = _b.sent();
                        result = siteResult && siteResult.value.length > 0 ? siteResult.value[0] : null;
                        _b.label = 7;
                    case 7: return [3 /*break*/, 9];
                    case 8: throw new Error("Unable to check site-scoped custom actions. ".concat(siteScopeResponse.status, " ").concat(siteScopeResponse.statusText));
                    case 9: return [3 /*break*/, 11];
                    case 10:
                        error_2 = _b.sent();
                        console.log("ERROR: Unable to fetch custom action with ClientSideComponentId ".concat(componentId, ". ").concat(error_2.message));
                        return [3 /*break*/, 11];
                    case 11: return [2 /*return*/, result];
                }
            });
        }); };
        this._context = context;
    }
    return ClientSideComponentService;
}());
export default ClientSideComponentService;
//# sourceMappingURL=ClientSideComponentService.js.map