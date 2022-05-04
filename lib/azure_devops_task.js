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
Object.defineProperty(exports, "__esModule", { value: true });
const tl = require("azure-pipelines-task-lib/task");
const store_apis_1 = require("./store_apis");
(function main() {
    return __awaiter(this, void 0, void 0, function* () {
        let storeApis = new store_apis_1.StoreApis();
        try {
            let command = tl.getInput("command");
            switch (command) {
                case "configure":
                    storeApis.productId = tl.getInput("productId");
                    storeApis.sellerId = tl.getInput("sellerId");
                    storeApis.tenantId = tl.getInput("tenantId");
                    storeApis.clientId = tl.getInput("clientId");
                    storeApis.clientSecret = tl.getInput("clientSecret");
                    yield storeApis.InitAsync();
                    tl.setVariable(`${store_apis_1.EnvVariablePrefix}product_id`, storeApis.productId, true, true);
                    tl.setVariable(`${store_apis_1.EnvVariablePrefix}seller_id`, storeApis.sellerId, true, true);
                    tl.setVariable(`${store_apis_1.EnvVariablePrefix}tenant_id`, storeApis.tenantId, true, true);
                    tl.setVariable(`${store_apis_1.EnvVariablePrefix}client_id`, storeApis.clientId, true, true);
                    tl.setVariable(`${store_apis_1.EnvVariablePrefix}client_secret`, storeApis.clientSecret, true, true);
                    tl.setVariable(`${store_apis_1.EnvVariablePrefix}access_token`, storeApis.accessToken, true, true);
                case "get":
                    let moduleName = tl.getInput("moduleName") || "";
                    let listingLanguage = tl.getInput("listingLanguage");
                    let draftSubmission = yield storeApis.GetExistingDraft(moduleName, listingLanguage);
                    tl.setVariable("draftSubmission", draftSubmission.toString());
                case "update":
                    let updatedProductString = tl.getInput("productUpdate");
                    if (!updatedProductString) {
                        tl.setResult(tl.TaskResult.Failed, `productUpdate parameter cannot be empty.`);
                        return;
                    }
                    let updateSubmissionData = yield storeApis.UpdateProductPackages(updatedProductString);
                    console.log(updateSubmissionData);
                case "poll":
                    let pollingSubmissionId = tl.getInput("pollingSubmissionId");
                    if (!pollingSubmissionId) {
                        tl.setResult(tl.TaskResult.Failed, `pollingSubmissionId parameter cannot be empty.`);
                        return;
                    }
                    let publishingStatus = yield storeApis.PollSubmissionStatus(pollingSubmissionId);
                    tl.setVariable("submissionStatus", publishingStatus);
                case "publish":
                    let submissionId = yield storeApis.PublishSubmission();
                    tl.setVariable("pollingSubmissionId", submissionId);
                default:
                    tl.setResult(tl.TaskResult.Failed, `Unknown command - ("${command}").`);
            }
        }
        catch (error) {
            tl.setResult(tl.TaskResult.Failed, error);
        }
    });
})();
//# sourceMappingURL=azure_devops_task.js.map