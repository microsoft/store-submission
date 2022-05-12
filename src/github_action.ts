import * as core from "@actions/core";
import { StoreApis, EnvVariablePrefix } from "./store_apis";

(async function main() {
  const storeApis = new StoreApis();

  try {
    const command = core.getInput("command");
    switch (command) {
      case "configure": {
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
      }

      case "get": {
        const moduleName = core.getInput("module-name");
        const listingLanguage = core.getInput("listing-language");
        const draftSubmission = await storeApis.GetExistingDraft(
          moduleName,
          listingLanguage
        );
        core.setOutput("draft-submission", draftSubmission);

        break;
      }

      case "update": {
        const updatedProductString = core.getInput("product-update");
        if (!updatedProductString) {
          core.setFailed(`product-update parameter cannot be empty.`);
          return;
        }

        const updateSubmissionData = await storeApis.UpdateProductPackages(
          updatedProductString
        );
        console.log(updateSubmissionData);

        break;
      }

      case "poll": {
        const pollingSubmissionId = core.getInput("polling-submission-id");

        if (!pollingSubmissionId) {
          core.setFailed(`polling-submission-id parameter cannot be empty.`);
          return;
        }

        const publishingStatus = await storeApis.PollSubmissionStatus(
          pollingSubmissionId
        );
        core.setOutput("submission-status", publishingStatus);

        break;
      }

      case "publish": {
        const submissionId = await storeApis.PublishSubmission();
        core.setOutput("polling-submission-id", submissionId);

        break;
      }

      default: {
        core.setFailed(`Unknown command - ("${command}").`);

        break;
      }
    }
  } catch (error: unknown) {
    core.setFailed(error as string);
  }
})();
