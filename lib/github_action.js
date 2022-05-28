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
const core = __importStar(require("@actions/core"));
const store_apis_1 = require("./store_apis");
(function main() {
    return __awaiter(this, void 0, void 0, function* () {
        const storeApis = new store_apis_1.StoreApis();
        try {
            const command = core.getInput("command");
            switch (command) {
                case "configure": {
                    storeApis.productId = core.getInput("product-id");
                    storeApis.sellerId = core.getInput("seller-id");
                    storeApis.tenantId = core.getInput("tenant-id");
                    storeApis.clientId = core.getInput("client-id");
                    storeApis.clientSecret = core.getInput("client-secret");
                    storeApis.onlyOnReady = core.getBooleanInput("only-on-ready");
                    yield storeApis.InitAsync();
                    core.exportVariable(`${store_apis_1.EnvVariablePrefix}product_id`, storeApis.productId);
                    core.exportVariable(`${store_apis_1.EnvVariablePrefix}seller_id`, storeApis.sellerId);
                    core.exportVariable(`${store_apis_1.EnvVariablePrefix}tenant_id`, storeApis.tenantId);
                    core.exportVariable(`${store_apis_1.EnvVariablePrefix}client_id`, storeApis.clientId);
                    core.exportVariable(`${store_apis_1.EnvVariablePrefix}client_secret`, storeApis.clientSecret);
                    core.exportVariable(`${store_apis_1.EnvVariablePrefix}access_token`, storeApis.accessToken);
                    core.exportVariable(`${store_apis_1.EnvVariablePrefix}only-on-ready`, storeApis.onlyOnReady);
                    core.setSecret(storeApis.productId);
                    core.setSecret(storeApis.sellerId);
                    core.setSecret(storeApis.tenantId);
                    core.setSecret(storeApis.clientId);
                    core.setSecret(storeApis.clientSecret);
                    core.setSecret(storeApis.accessToken);
                    break;
                }
                case "get": {
                    const moduleName = core.getInput("module-name");
                    const listingLanguage = core.getInput("listing-language");
                    const draftSubmission = yield storeApis.GetExistingDraft(moduleName, listingLanguage);
                    core.setOutput("draft-submission", draftSubmission);
                    break;
                }
                case "update": {
                    if (!(yield storeApis.IsReady())) {
                        core.notice(`Only on ready is set and module is not ready, skipping.`);
                        return;
                    }
                    const updatedMetadataString = core.getInput("metadata-update");
                    const updatedProductString = core.getInput("product-update");
                    if (!updatedMetadataString && !updatedProductString) {
                        core.setFailed(`Nothing to update. Both product-update and metadata-update are null.`);
                        return;
                    }
                    if (updatedMetadataString) {
                        const updateSubmissionMetadata = yield storeApis.UpdateSubmissionMetadata(updatedMetadataString);
                        console.log(updateSubmissionMetadata);
                    }
                    if (updatedProductString) {
                        const updateSubmissionData = yield storeApis.UpdateProductPackages(updatedProductString);
                        console.log(updateSubmissionData);
                    }
                    break;
                }
                case "poll": {
                    const pollingSubmissionId = core.getInput("polling-submission-id");
                    if (!pollingSubmissionId) {
                        core.setFailed(`polling-submission-id parameter cannot be empty.`);
                        return;
                    }
                    const publishingStatus = yield storeApis.PollSubmissionStatus(pollingSubmissionId);
                    core.setOutput("submission-status", publishingStatus);
                    break;
                }
                case "publish": {
                    if (!(yield storeApis.IsReady())) {
                        core.notice(`Only on ready is set and module is not ready, skipping`);
                        return;
                    }
                    const submissionId = yield storeApis.PublishSubmission();
                    core.setOutput("polling-submission-id", submissionId);
                    break;
                }
                default: {
                    core.setFailed(`Unknown command - ("${command}").`);
                    break;
                }
            }
        }
        catch (error) {
            core.setFailed(error);
        }
    });
})();
//# sourceMappingURL=github_action.js.map