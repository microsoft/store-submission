import * as core from "@actions/core";
import { StoreApis, EnvVariablePrefix } from "./store_apis";

(async function main() {
  let storeApis = new StoreApis();

  try {
    let command = core.getInput("command");
    switch (command) {
      case "configure":
        storeApis.productId = core.getInput("product-id");
        storeApis.sellerId = core.getInput("seller-id");
        storeApis.tenantId = core.getInput("tenant-id");
        storeApis.clientId = core.getInput("client-id");
        storeApis.clientSecret = core.getInput("client-secret");

        await storeApis.InitAsync();

        core.exportVariable(
          `${EnvVariablePrefix}product_id`,
          storeApis.productId
        );
        core.exportVariable(
          `${EnvVariablePrefix}seller_id`,
          storeApis.sellerId
        );
        core.exportVariable(
          `${EnvVariablePrefix}tenant_id`,
          storeApis.tenantId
        );
        core.exportVariable(
          `${EnvVariablePrefix}client_id`,
          storeApis.clientId
        );
        core.exportVariable(
          `${EnvVariablePrefix}client_secret`,
          storeApis.clientSecret
        );
        core.exportVariable(
          `${EnvVariablePrefix}access_token`,
          storeApis.accessToken
        );
        core.setSecret(storeApis.productId);
        core.setSecret(storeApis.sellerId);
        core.setSecret(storeApis.tenantId);
        core.setSecret(storeApis.clientId);
        core.setSecret(storeApis.clientSecret);
        core.setSecret(storeApis.accessToken);

        break;

      case "get":
        let moduleName = core.getInput("module-name");
        let listingLanguage = core.getInput("listing-language");
        let draftSubmission = await storeApis.GetExistingDraft(
          moduleName,
          listingLanguage
        );
        core.setOutput("draft-submission", draftSubmission);

        break;

      case "update":
        let updatedProductString = core.getInput("product-update");
        if (!updatedProductString) {
          core.setFailed(`product-update parameter cannot be empty.`);
          return;
        }

        let updateSubmissionData = await storeApis.UpdateProductPackages(
          updatedProductString
        );
        console.log(updateSubmissionData);

        break;

      case "poll":
        let pollingSubmissionId = core.getInput("polling-submission-id");

        if (!pollingSubmissionId) {
          core.setFailed(`polling-submission-id parameter cannot be empty.`);
          return;
        }

        let publishingStatus = await storeApis.PollSubmissionStatus(
          pollingSubmissionId
        );
        core.setOutput("submission-status", publishingStatus);

        break;

      case "publish":
        let submissionId = await storeApis.PublishSubmission();
        core.setOutput("polling-submission-id", submissionId);

        break;

      default:
        core.setFailed(`Unknown command - ("${command}").`);

        break;
    }
  } catch (error: any) {
    core.setFailed(error);
  }
})();
