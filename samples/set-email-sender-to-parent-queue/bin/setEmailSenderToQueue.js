"use strict";
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
exports.__esModule = true;
exports.setEmailSenderToQueue = void 0;
// Export function to make it available for import in other files
function setEmailSenderToQueue(ctx) {
    return __awaiter(this, void 0, void 0, function () {
        var formContext, regarding, queue;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    formContext = ctx.getFormContext();
                    if (!(formContext.ui.getFormType() === 1 /* Create */ || formContext.getAttribute("from").getValue() === null)) return [3 /*break*/, 2];
                    regarding = formContext.getAttribute("regardingobjectid").getValue();
                    if (!(regarding !== null && regarding[0].entityType.toLowerCase() === "incident")) return [3 /*break*/, 2];
                    return [4 /*yield*/, Xrm.WebApi.retrieveMultipleRecords("queueitem", "?$select=_queueid_value&$filter=statecode eq 0 and _objectid_value eq " + regarding[0].id).then(getEntityReferenceFromMultipleResult)];
                case 1:
                    queue = _a.sent();
                    // Set sender to queue if exists
                    if (queue !== null) {
                        formContext.getAttribute("from").setValue([queue]);
                    }
                    _a.label = 2;
                case 2: return [2 /*return*/];
            }
        });
    });
}
exports.setEmailSenderToQueue = setEmailSenderToQueue;
function getEntityReferenceFromMultipleResult(results) {
    return __awaiter(this, void 0, void 0, function () {
        var queue;
        return __generator(this, function (_a) {
            // Return lookup value of queue if exactly 1 active queueitem is found
            if (results.entities.length === 1) {
                queue = {};
                queue.entityType = results.entities[0]["_queueid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
                queue.id = results.entities[0]._queueid_value;
                queue.name = results.entities[0]["_queueid_value@OData.Community.Display.V1.FormattedValue"];
                return [2 /*return*/, queue];
            }
            // Throw error if more than 1 active queueitem is found
            if (results.entities.length > 1) {
                throw "Number of active queue items is more than 1";
            }
            // Return null by default
            return [2 /*return*/, null];
        });
    });
}
