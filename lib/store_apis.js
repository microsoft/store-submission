"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.StoreApis = exports.EnvVariablePrefix = void 0;
const https = __importStar(require("https"));
exports.EnvVariablePrefix = "MICROSOFT_STORE_ACTION_";
class ResponseWrapper {
    constructor() {
        Object.defineProperty(this, "responseData", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "isSuccess", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "errors", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class ErrorResponse {
    constructor() {
        Object.defineProperty(this, "statusCode", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "message", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class Error {
    constructor() {
        Object.defineProperty(this, "code", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "message", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "target", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class SubmissionResponse {
    constructor() {
        Object.defineProperty(this, "pollingUrl", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "submissionId", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "ongoingSubmissionId", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
var PublishingStatus;
(function (PublishingStatus) {
    PublishingStatus["INPROGRESS"] = "INPROGRESS";
    PublishingStatus["PUBLISHED"] = "PUBLISHED";
    PublishingStatus["FAILED"] = "FAILED";
    PublishingStatus["UNKNOWN"] = "UNKNOWN";
})(PublishingStatus || (PublishingStatus = {}));
class ModuleStatus {
    constructor() {
        Object.defineProperty(this, "isReady", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "ongoingSubmissionId", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class SubmissionStatus {
    constructor() {
        Object.defineProperty(this, "publishingStatus", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "hasFailed", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class ListingAssetsResponse {
    constructor() {
        Object.defineProperty(this, "listingAssets", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class ImageSize {
    constructor() {
        Object.defineProperty(this, "width", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "height", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class ListingAsset {
    constructor() {
        Object.defineProperty(this, "language", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "storeLogos", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "screenshots", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class Screenshot {
    constructor() {
        Object.defineProperty(this, "id", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "assetUrl", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "imageSize", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class StoreLogo {
    constructor() {
        Object.defineProperty(this, "id", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "assetUrl", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "imageSize", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
    }
}
class StoreApis {
    constructor() {
        Object.defineProperty(this, "accessToken", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "productId", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "sellerId", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "tenantId", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "clientId", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "clientSecret", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        Object.defineProperty(this, "onlyOnReady", {
            enumerable: true,
            configurable: true,
            writable: true,
            value: void 0
        });
        this.LoadState();
    }
    Delay(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }
    GetAccessToken() {
        return __awaiter(this, void 0, void 0, function* () {
            const requestParameters = {
                grant_type: "client_credentials",
                client_id: this.clientId,
                client_secret: this.clientSecret,
                scope: StoreApis.scope,
            };
            const formBody = [];
            for (const property in requestParameters) {
                const encodedKey = encodeURIComponent(property);
                const encodedValue = encodeURIComponent(requestParameters[property]);
                formBody.push(encodedKey + "=" + encodedValue);
            }
            const dataString = formBody.join("\r\n&");
            const options = {
                host: StoreApis.microsoftOnlineLoginHost,
                path: `/${this.tenantId}${StoreApis.authOAuth2TokenSuffix}`,
                method: "POST",
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Content-Length": dataString.length,
                },
            };
            return new Promise((resolve, reject) => {
                const req = https.request(options, (res) => {
                    let responseString = "";
                    res.on("data", (data) => {
                        responseString += data;
                    });
                    res.on("end", function () {
                        const responseObject = JSON.parse(responseString);
                        if (responseObject.error)
                            reject(responseObject);
                        else
                            resolve(responseObject.access_token);
                    });
                });
                req.on("error", (e) => {
                    console.error(e);
                    reject(e);
                });
                req.write(dataString);
                req.end();
            });
        });
    }
    GetCurrentDraftSubmissionPackagesData() {
        return this.CreateStoreHttpRequest("", "GET", `/submission/v1/product/${this.productId}/packages`);
    }
    GetCurrentDraftSubmissionMetadata(moduleName, listingLanguages) {
        return this.CreateStoreHttpRequest("", "GET", `/submission/v1/product/${this.productId}/metadata/${moduleName}?languages=${listingLanguages}`);
    }
    UpdateCurrentDraftSubmissionMetadata(submissionMetadata) {
        return this.CreateStoreHttpRequest(JSON.stringify(submissionMetadata), "PUT", `/submission/v1/product/${this.productId}/metadata`);
    }
    UpdateStoreSubmissionPackages(submission) {
        return __awaiter(this, void 0, void 0, function* () {
            return this.CreateStoreHttpRequest(JSON.stringify(submission), "PUT", `/submission/v1/product/${this.productId}/packages`);
        });
    }
    CommitUpdateStoreSubmissionPackages() {
        return __awaiter(this, void 0, void 0, function* () {
            return this.CreateStoreHttpRequest("", "POST", `/submission/v1/product/${this.productId}/packages/commit`);
        });
    }
    GetModuleStatus() {
        return __awaiter(this, void 0, void 0, function* () {
            return this.CreateStoreHttpRequest("", "GET", `/submission/v1/product/${this.productId}/status`);
        });
    }
    GetSubmissionStatus(submissionId) {
        return __awaiter(this, void 0, void 0, function* () {
            return this.CreateStoreHttpRequest("", "GET", `/submission/v1/product/${this.productId}/submission/${submissionId}/status`);
        });
    }
    SubmitSubmission() {
        return __awaiter(this, void 0, void 0, function* () {
            return this.CreateStoreHttpRequest("", "POST", `/submission/v1/product/${this.productId}/submit`);
        });
    }
    GetCurrentDraftListingAssets(listingLanguages) {
        return __awaiter(this, void 0, void 0, function* () {
            return this.CreateStoreHttpRequest("", "GET", `/submission/v1/product/${this.productId}/listings/assets?languages=${listingLanguages}`);
        });
    }
    CreateStoreHttpRequest(requestParameters, method, path) {
        return __awaiter(this, void 0, void 0, function* () {
            const options = {
                host: StoreApis.storeApiUrl,
                path: path,
                method: method,
                headers: {
                    "Content-Type": "application/json",
                    "Content-Length": requestParameters.length,
                    Authorization: "Bearer " + this.accessToken,
                    "X-Seller-Account-Id": this.sellerId,
                },
            };
            return new Promise((resolve, reject) => {
                const req = https.request(options, (res) => {
                    if (res.statusCode == 404) {
                        const error = new ResponseWrapper();
                        error.isSuccess = false;
                        error.errors = [];
                        error.errors[0] = new Error();
                        error.errors[0].message = "Not found";
                        reject(error);
                        return;
                    }
                    let responseString = "";
                    res.on("data", (data) => {
                        responseString += data;
                    });
                    res.on("end", function () {
                        const responseObject = JSON.parse(responseString);
                        resolve(responseObject);
                    });
                });
                req.on("error", (e) => {
                    console.error(e);
                    reject(e);
                });
                req.write(requestParameters);
                req.end();
            });
        });
    }
    PollModuleStatus() {
        return __awaiter(this, void 0, void 0, function* () {
            let status = new ModuleStatus();
            status.isReady = false;
            while (!status.isReady) {
                const moduleStatus = yield this.GetModuleStatus();
                console.log(JSON.stringify(moduleStatus));
                status = moduleStatus.responseData;
                if (!moduleStatus.isSuccess) {
                    const errorResponse = moduleStatus;
                    if (errorResponse.statusCode == 401) {
                        console.log(`Access token expired. Requesting new one. (message='${errorResponse.message}')`);
                        yield this.InitAsync();
                        status = new ModuleStatus();
                        status.isReady = false;
                        continue;
                    }
                    console.log("Error");
                    break;
                }
                if (status.isReady) {
                    console.log("Success!");
                    return true;
                }
                else {
                    if (moduleStatus.errors &&
                        moduleStatus.errors.length > 0 &&
                        moduleStatus.errors.find((e) => e.target != "packages" || e.code == "packageuploaderror")) {
                        console.log(moduleStatus.errors);
                        return false;
                    }
                }
                console.log("Waiting 10 seconds.");
                yield this.Delay(10000);
            }
            return false;
        });
    }
    InitAsync() {
        return __awaiter(this, void 0, void 0, function* () {
            this.accessToken = yield this.GetAccessToken();
        });
    }
    IsReady() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.onlyOnReady) {
                return true;
            }
            const moduleStatus = yield this.GetModuleStatus();
            return moduleStatus.responseData.isReady;
        });
    }
    GetExistingDraft(moduleName, listingLanguage) {
        return __awaiter(this, void 0, void 0, function* () {
            return new Promise((resolve, reject) => {
                if (moduleName &&
                    moduleName.toLowerCase() != "availability" &&
                    moduleName.toLowerCase() != "listings" &&
                    moduleName.toLowerCase() != "properties") {
                    reject("Module name must be 'availability', 'listings' or 'properties'");
                    return;
                }
                (moduleName
                    ? this.GetCurrentDraftSubmissionMetadata(moduleName, listingLanguage)
                    : this.GetCurrentDraftSubmissionPackagesData())
                    .then((currentDraftResponse) => {
                    if (!currentDraftResponse.isSuccess) {
                        reject(`Failed to get the existing draft. - ${JSON.stringify(currentDraftResponse, null, 2)}`);
                    }
                    else {
                        resolve(JSON.stringify(currentDraftResponse.responseData));
                    }
                })
                    .catch((error) => {
                    reject(`Failed to get the existing draft. - ${error.errors}`);
                });
            });
        });
    }
    PollSubmissionStatus(pollingSubmissionId) {
        return __awaiter(this, void 0, void 0, function* () {
            let status = new SubmissionStatus();
            status.hasFailed = false;
            // eslint-disable-next-line no-async-promise-executor
            return new Promise((resolve, reject) => __awaiter(this, void 0, void 0, function* () {
                while (!status.hasFailed) {
                    const submissionStatus = yield this.GetSubmissionStatus(pollingSubmissionId);
                    console.log(JSON.stringify(submissionStatus));
                    status = submissionStatus.responseData;
                    if (!submissionStatus.isSuccess || status.hasFailed) {
                        const errorResponse = submissionStatus;
                        if (errorResponse.statusCode == 401) {
                            console.log(`Access token expired. Requesting new one. (message='${errorResponse.message}')`);
                            yield this.InitAsync();
                            status = new SubmissionStatus();
                            status.hasFailed = false;
                            continue;
                        }
                        console.log("Error");
                        reject("Error");
                        return;
                    }
                    if (status.publishingStatus == PublishingStatus.PUBLISHED ||
                        status.publishingStatus == PublishingStatus.FAILED) {
                        console.log(`PublishingStatus = ${status.publishingStatus}`);
                        resolve(status.publishingStatus);
                        return;
                    }
                    console.log("Waiting 10 seconds.");
                    yield this.Delay(10000);
                }
            }));
        });
    }
    UpdateSubmissionMetadata(submissionMetadataString) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!(yield this.PollModuleStatus())) {
                // Wait until all modules are in the ready state
                return Promise.reject("Failed to poll module status.");
            }
            const submissionMetadata = JSON.parse(submissionMetadataString);
            console.log(submissionMetadata);
            const updateSubmissionData = yield this.UpdateCurrentDraftSubmissionMetadata(submissionMetadata);
            console.log(JSON.stringify(updateSubmissionData));
            if (!updateSubmissionData.isSuccess) {
                return Promise.reject(`Failed to update submission metadata - ${JSON.stringify(updateSubmissionData.errors)}`);
            }
            if (!(yield this.PollModuleStatus())) {
                // Wait until all modules are in the ready state
                return Promise.reject("Failed to poll module status.");
            }
            return updateSubmissionData;
        });
    }
    UpdateProductPackages(updatedProductString) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!(yield this.PollModuleStatus())) {
                // Wait until all modules are in the ready state
                return Promise.reject("Failed to poll module status.");
            }
            const updatedProductPackages = JSON.parse(updatedProductString);
            console.log(updatedProductPackages);
            const updateSubmissionData = yield this.UpdateStoreSubmissionPackages(updatedProductPackages);
            console.log(JSON.stringify(updateSubmissionData));
            if (!updateSubmissionData.isSuccess) {
                return Promise.reject(`Failed to update submission - ${JSON.stringify(updateSubmissionData.errors)}`);
            }
            console.log("Committing package changes...");
            const commitResult = yield this.CommitUpdateStoreSubmissionPackages();
            if (!commitResult.isSuccess) {
                return Promise.reject(`Failed to commit the updated submission - ${JSON.stringify(commitResult.errors)}`);
            }
            console.log(JSON.stringify(commitResult));
            if (!(yield this.PollModuleStatus())) {
                // Wait until all modules are in the ready state
                return Promise.reject("Failed to poll module status.");
            }
            return updateSubmissionData;
        });
    }
    PublishSubmission() {
        return __awaiter(this, void 0, void 0, function* () {
            const commitResult = yield this.CommitUpdateStoreSubmissionPackages();
            if (!commitResult.isSuccess) {
                return Promise.reject(`Failed to commit the updated submission - ${JSON.stringify(commitResult.errors)}`);
            }
            console.log(JSON.stringify(commitResult));
            if (!(yield this.PollModuleStatus())) {
                // Wait until all modules are in the ready state
                return Promise.reject("Failed to poll module status.");
            }
            let submissionId = null;
            const submitSubmissionResponse = yield this.SubmitSubmission();
            console.log(JSON.stringify(submitSubmissionResponse));
            if (submitSubmissionResponse.isSuccess) {
                if (submitSubmissionResponse.responseData.submissionId != null &&
                    submitSubmissionResponse.responseData.submissionId.length > 0) {
                    submissionId = submitSubmissionResponse.responseData.submissionId;
                }
                else if (submitSubmissionResponse.responseData.ongoingSubmissionId != null &&
                    submitSubmissionResponse.responseData.ongoingSubmissionId.length > 0) {
                    submissionId =
                        submitSubmissionResponse.responseData.ongoingSubmissionId;
                }
            }
            return new Promise((resolve, reject) => {
                if (submissionId == null) {
                    console.log("Failed to get submission ID");
                    reject("Failed to get submission ID");
                }
                else {
                    resolve(submissionId);
                }
            });
        });
    }
    GetExistingDraftListingAssets(listingLanguage) {
        return __awaiter(this, void 0, void 0, function* () {
            return new Promise((resolve, reject) => {
                this.GetCurrentDraftListingAssets(listingLanguage)
                    .then((draftListingAssetsResponse) => {
                    if (!draftListingAssetsResponse.isSuccess) {
                        reject(`Failed to get the existing draft listing assets. - ${JSON.stringify(draftListingAssetsResponse, null, 2)}`);
                    }
                    else {
                        resolve(JSON.stringify(draftListingAssetsResponse.responseData));
                    }
                })
                    .catch((error) => {
                    reject(`Failed to get the existing draft listing assets. - ${error.errors}`);
                });
            });
        });
    }
    LoadState() {
        var _a, _b, _c, _d, _e, _f, _g;
        this.productId = (_a = process.env[`${exports.EnvVariablePrefix}product_id`]) !== null && _a !== void 0 ? _a : "";
        this.sellerId = (_b = process.env[`${exports.EnvVariablePrefix}seller_id`]) !== null && _b !== void 0 ? _b : "";
        this.tenantId = (_c = process.env[`${exports.EnvVariablePrefix}tenant_id`]) !== null && _c !== void 0 ? _c : "";
        this.clientId = (_d = process.env[`${exports.EnvVariablePrefix}client_id`]) !== null && _d !== void 0 ? _d : "";
        this.clientSecret = (_e = process.env[`${exports.EnvVariablePrefix}client_secret`]) !== null && _e !== void 0 ? _e : "";
        this.accessToken = (_f = process.env[`${exports.EnvVariablePrefix}access_token`]) !== null && _f !== void 0 ? _f : "";
        const onlyOnReady = (_g = process.env[`${exports.EnvVariablePrefix}only-on-ready`]) !== null && _g !== void 0 ? _g : "false";
        this.onlyOnReady = onlyOnReady === "true";
    }
}
exports.StoreApis = StoreApis;
Object.defineProperty(StoreApis, "microsoftOnlineLoginHost", {
    enumerable: true,
    configurable: true,
    writable: true,
    value: "login.microsoftonline.com"
});
Object.defineProperty(StoreApis, "authOAuth2TokenSuffix", {
    enumerable: true,
    configurable: true,
    writable: true,
    value: "/oauth2/v2.0/token"
});
Object.defineProperty(StoreApis, "scope", {
    enumerable: true,
    configurable: true,
    writable: true,
    value: "https://api.store.microsoft.com/.default"
});
Object.defineProperty(StoreApis, "storeApiUrl", {
    enumerable: true,
    configurable: true,
    writable: true,
    value: "api.store.microsoft.com"
});
//# sourceMappingURL=store_apis.js.map