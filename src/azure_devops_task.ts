import tl = require("azure-pipelines-task-lib/task");
import { StoreApis, EnvVariablePrefix } from "./store_apis";

(async function main() {
  let storeApis = new StoreApis();

  try {
    let command = tl.getInput("command");
    switch (command) {
      case "configure":
        storeApis.productId = tl.getInput("productId")!;
        storeApis.sellerId = tl.getInput("sellerId")!;
        storeApis.tenantId = tl.getInput("tenantId")!;
        storeApis.clientId = tl.getInput("clientId")!;
        storeApis.clientSecret = tl.getInput("clientSecret")!;

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

      case "get":
        let moduleName = tl.getInput("moduleName") || "";
        let listingLanguage = tl.getInput("listingLanguage")!;
        let draftSubmission = await storeApis.GetExistingDraft(
          moduleName,
          listingLanguage
        );
        tl.setVariable("draftSubmission", draftSubmission.toString());

      case "update":
        let updatedProductString = tl.getInput("productUpdate");
        if (!updatedProductString) {
          tl.setResult(
            tl.TaskResult.Failed,
            `productUpdate parameter cannot be empty.`
          );
          return;
        }

        let updateSubmissionData = await storeApis.UpdateProductPackages(
          updatedProductString
        );
        console.log(updateSubmissionData);

      case "poll":
        let pollingSubmissionId = tl.getInput("pollingSubmissionId");

        if (!pollingSubmissionId) {
          tl.setResult(
            tl.TaskResult.Failed,
            `pollingSubmissionId parameter cannot be empty.`
          );
          return;
        }

        let publishingStatus = await storeApis.PollSubmissionStatus(
          pollingSubmissionId
        );
        tl.setVariable("submissionStatus", publishingStatus);

      case "publish":
        let submissionId = await storeApis.PublishSubmission();
        tl.setVariable("pollingSubmissionId", submissionId);

      default:
        tl.setResult(tl.TaskResult.Failed, `Unknown command - ("${command}").`);
    }
  } catch (error: any) {
    tl.setResult(tl.TaskResult.Failed, error);
  }
})();
