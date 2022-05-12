import tl = require("azure-pipelines-task-lib/task");
import { StoreApis, EnvVariablePrefix } from "./store_apis";

(async function main() {
  const storeApis = new StoreApis();

  try {
    const command = tl.getInput("command");
    switch (command) {
      case "configure": {
        storeApis.productId = tl.getInput("productId") || "";
        storeApis.sellerId = tl.getInput("sellerId") || "";
        storeApis.tenantId = tl.getInput("tenantId") || "";
        storeApis.clientId = tl.getInput("clientId") || "";
        storeApis.clientSecret = tl.getInput("clientSecret") || "";

        await storeApis.InitAsync();

        tl.setVariable(
          `${EnvVariablePrefix}product_id`,
          storeApis.productId,
          true,
          true
        );
        tl.setVariable(
          `${EnvVariablePrefix}seller_id`,
          storeApis.sellerId,
          true,
          true
        );
        tl.setVariable(
          `${EnvVariablePrefix}tenant_id`,
          storeApis.tenantId,
          true,
          true
        );
        tl.setVariable(
          `${EnvVariablePrefix}client_id`,
          storeApis.clientId,
          true,
          true
        );
        tl.setVariable(
          `${EnvVariablePrefix}client_secret`,
          storeApis.clientSecret,
          true,
          true
        );
        tl.setVariable(
          `${EnvVariablePrefix}access_token`,
          storeApis.accessToken,
          true,
          true
        );

        break;
      }

      case "get": {
        const moduleName = tl.getInput("moduleName") || "";
        const listingLanguage = tl.getInput("listingLanguage") || "en";
        const draftSubmission = await storeApis.GetExistingDraft(
          moduleName,
          listingLanguage
        );
        tl.setVariable("draftSubmission", draftSubmission.toString());

        break;
      }

      case "update": {
        const updatedProductString = tl.getInput("productUpdate");
        if (!updatedProductString) {
          tl.setResult(
            tl.TaskResult.Failed,
            `productUpdate parameter cannot be empty.`
          );
          return;
        }

        const updateSubmissionData = await storeApis.UpdateProductPackages(
          updatedProductString
        );
        console.log(updateSubmissionData);

        break;
      }

      case "poll": {
        const pollingSubmissionId = tl.getInput("pollingSubmissionId");

        if (!pollingSubmissionId) {
          tl.setResult(
            tl.TaskResult.Failed,
            `pollingSubmissionId parameter cannot be empty.`
          );
          return;
        }

        const publishingStatus = await storeApis.PollSubmissionStatus(
          pollingSubmissionId
        );
        tl.setVariable("submissionStatus", publishingStatus);

        break;
      }

      case "publish": {
        const submissionId = await storeApis.PublishSubmission();
        tl.setVariable("pollingSubmissionId", submissionId);

        break;
      }

      default: {
        tl.setResult(tl.TaskResult.Failed, `Unknown command - ("${command}").`);

        break;
      }
    }
  } catch (error: unknown) {
    tl.setResult(tl.TaskResult.Failed, error as string);
  }
})();
