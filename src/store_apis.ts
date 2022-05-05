import * as https from "https";

export let EnvVariablePrefix = "MICROSOFT_STORE_ACTION_";

class ResponseWrapper<T> {
  responseData: T;
  isSuccess: boolean;
  errors: Error[];
}

class ErrorResponse {
  statusCode: number;
  message: string;
}

class Error {
  code: string;
  message: string;
  target: string;
}

class SubmissionResponse {
  pollingUrl: string;
  submissionId: string;
  ongoingSubmissionId: string;
}

enum PublishingStatus {
  INPROGRESS = "INPROGRESS",
  PUBLISHED = "PUBLISHED",
  FAILED = "FAILED",
  UNKNOWN = "UNKNOWN",
}

class ModuleStatus {
  isReady: boolean;
  ongoingSubmissionId: string;
}

class SubmissionStatus {
  publishingStatus: PublishingStatus;
  hasFailed: boolean;
}

export class StoreApis {
  private static readonly microsoftOnlineLoginHost =
    "login.microsoftonline.com";
  private static readonly authOAuth2TokenSuffix = "/oauth2/v2.0/token";
  private static readonly scope =
    "https://api.store-int.microsoft.com/.default";
  private static readonly storeApiUrl = "api.store-int.microsoft.com";
  // private static readonly storeApiUrl = 'api.store.microsoft.com'

  constructor() {
    this.LoadState();
  }

  public accessToken: string;
  public productId: string;
  public sellerId: string;
  public tenantId: string;
  public clientId: string;
  public clientSecret: string;

  private Delay(ms: number): Promise<unknown> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private async GetAccessToken(): Promise<string> {
    const requestParameters: { [key: string]: string } = {
      grant_type: "client_credentials",
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: StoreApis.scope,
    };

    var formBody = [];
    for (var property in requestParameters) {
      var encodedKey = encodeURIComponent(property);
      var encodedValue = encodeURIComponent(requestParameters[property]);
      formBody.push(encodedKey + "=" + encodedValue);
    }

    var dataString = formBody.join("\r\n&");

    var options: https.RequestOptions = {
      host: StoreApis.microsoftOnlineLoginHost,
      path: `/${this.tenantId}${StoreApis.authOAuth2TokenSuffix}`,
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Content-Length": dataString.length,
      },
    };

    return new Promise<string>((resolve, reject) => {
      var req = https.request(options, (res) => {
        var responseString: string = "";

        res.on("data", (data) => {
          responseString += data;
        });

        res.on("end", function () {
          var responseObject = JSON.parse(responseString);
          if (responseObject.error) reject(responseObject);
          else resolve(responseObject.access_token);
        });
      });

      req.on("error", (e) => {
        console.error(e);
        reject(e);
      });

      req.write(dataString);
      req.end();
    });
  }

  public GetCurrentDraftSubmissionPackagesData(): Promise<
    ResponseWrapper<any>
  > {
    return this.CreateStoreHttpRequest(
      "",
      "GET",
      `/submission/v1/product/${this.productId}/packages`
    );
  }

  public GetCurrentDraftSubmissionMetadata(
    moduleName: string,
    listingLanguages: string
  ): Promise<ResponseWrapper<any>> {
    return this.CreateStoreHttpRequest(
      "",
      "GET",
      `/submission/v1/product/${this.productId}/metadata?languages=${listingLanguages}`
    );
  }

  public async UpdateStoreSubmissionPackages(
    submission: any
  ): Promise<ResponseWrapper<any>> {
    return this.CreateStoreHttpRequest(
      JSON.stringify(submission),
      "PUT",
      `/submission/v1/product/${this.productId}/packages`
    );
  }

  public async CommitUpdateStoreSubmissionPackages(): Promise<
    ResponseWrapper<any>
  > {
    return this.CreateStoreHttpRequest(
      "",
      "POST",
      `/submission/v1/product/${this.productId}/packages/commit`
    );
  }

  public async GetModuleStatus(): Promise<ResponseWrapper<ModuleStatus>> {
    return this.CreateStoreHttpRequest<ModuleStatus>(
      "",
      "GET",
      `/submission/v1/product/${this.productId}/status`
    );
  }

  public async GetSubmissionStatus(
    submissionId: string
  ): Promise<ResponseWrapper<SubmissionStatus>> {
    return this.CreateStoreHttpRequest<SubmissionStatus>(
      "",
      "GET",
      `/submission/v1/product/${this.productId}/submission/${submissionId}/status`
    );
  }

  public async SubmitSubmission(): Promise<
    ResponseWrapper<SubmissionResponse>
  > {
    return this.CreateStoreHttpRequest<SubmissionResponse>(
      "",
      "POST",
      `/submission/v1/product/${this.productId}/submit`
    );
  }

  public async CreateStoreHttpRequest<T>(
    requestParameters: string,
    method: string,
    path: string
  ): Promise<ResponseWrapper<T>> {
    var options: https.RequestOptions = {
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

    return new Promise<ResponseWrapper<T>>((resolve, reject) => {
      var req = https.request(options, (res) => {
        if (res.statusCode == 404) {
          let error = new ResponseWrapper();
          error.isSuccess = false;
          error.errors = [];
          error.errors[0] = new Error();
          error.errors[0].message = "Not found";
          reject(error);
          return;
        }
        var responseString: string = "";

        res.on("data", (data) => {
          responseString += data;
        });

        res.on("end", function () {
          var responseObject = <ResponseWrapper<T>>JSON.parse(responseString);
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
  }

  public async PollModuleStatus(): Promise<boolean> {
    let status: ModuleStatus = new ModuleStatus();
    status.isReady = false;

    while (!status.isReady) {
      console.log("Waiting 10 seconds.");
      await this.Delay(10000);
      let moduleStatus = await this.GetModuleStatus();
      console.log(JSON.stringify(moduleStatus));
      status = moduleStatus.responseData;
      if (!moduleStatus.isSuccess) {
        var errorResponse = moduleStatus as any as ErrorResponse;
        if (errorResponse.statusCode == 401) {
          console.log(
            `Access token expired. Requesting new one. (message='${errorResponse.message}')`
          );
          await this.InitAsync();
          status.isReady = false;
          continue;
        }
        console.log("Error");
        break;
      }
      if (status.isReady) {
        console.log("Success!");
        return true;
      } else {
        if (
          moduleStatus.errors &&
          moduleStatus.errors.length > 0 &&
          moduleStatus.errors.find(
            (e) => e.target != "packages" || e.code == "packageuploaderror"
          )
        ) {
          console.log(moduleStatus.errors);
          return false;
        }
      }
    }

    return false;
  }

  public async InitAsync() {
    this.accessToken = await this.GetAccessToken();
  }

  public async GetExistingDraft(
    moduleName: string,
    listingLanguage: string
  ): Promise<string> {
    return new Promise<string>(async (resolve, reject) => {
      await (moduleName
        ? this.GetCurrentDraftSubmissionMetadata(moduleName, listingLanguage)
        : this.GetCurrentDraftSubmissionPackagesData()
      )
        .then((currentDraftResponse) => {
          if (!currentDraftResponse.isSuccess) {
            reject("Failed to get the existing draft.");
          } else {
            resolve(JSON.stringify(currentDraftResponse.responseData));
          }
        })
        .catch((error: any) => {
          reject(`Failed to get the existing draft. - ${error.errorS}`);
        });
    });
  }

  public async PollSubmissionStatus(
    pollingSubmissionId: string
  ): Promise<PublishingStatus> {
    let status: SubmissionStatus = new SubmissionStatus();
    status.hasFailed = false;

    return new Promise<PublishingStatus>(async (resolve, reject) => {
      while (!status.hasFailed) {
        let submissionStatus = await this.GetSubmissionStatus(
          pollingSubmissionId
        );
        console.log(JSON.stringify(submissionStatus));
        status = submissionStatus.responseData;
        if (!submissionStatus.isSuccess || status.hasFailed) {
          var errorResponse = submissionStatus as any as ErrorResponse;
          if (errorResponse.statusCode == 401) {
            console.log(
              `Access token expired. Requesting new one. (message='${errorResponse.message}')`
            );
            await this.InitAsync();
            status = new SubmissionStatus();
            status.hasFailed = false;
            continue;
          }
          console.log("Error");
          reject("Error");
          return;
        }
        if (
          status.publishingStatus == PublishingStatus.PUBLISHED ||
          status.publishingStatus == PublishingStatus.FAILED
        ) {
          console.log(`PublishingStatus = ${status.publishingStatus}`);
          resolve(status.publishingStatus);
          return;
        }

        console.log("Waiting 10 seconds.");
        await this.Delay(10000);
      }
    });
  }

  public async UpdateProductPackages(
    updatedProductString: string
  ): Promise<any> {
    let updatedProductPackages = JSON.parse(updatedProductString);

    console.log(updatedProductPackages);

    let updateSubmissionData = await this.UpdateStoreSubmissionPackages(
      updatedProductPackages
    );
    console.log(JSON.stringify(updateSubmissionData));

    if (!updateSubmissionData.isSuccess) {
      return Promise.reject(
        `Failed to update submission - ${updateSubmissionData.errors}`
      );
    }

    let commitResult = await this.CommitUpdateStoreSubmissionPackages();
    if (!commitResult.isSuccess) {
      return Promise.reject(
        `Failed to commit the updated submission - ${commitResult.errors}`
      );
    }

    return updateSubmissionData;
  }

  public async PublishSubmission(): Promise<string> {
    if (!(await this.PollModuleStatus())) {
      // Wait until all modules are in the ready state
      return Promise.reject("Failed to poll module status.");
    }

    let submissionId: string | null = null;

    let submitSubmissionResponse: ResponseWrapper<SubmissionResponse>;
    submitSubmissionResponse = await this.SubmitSubmission();
    console.log(JSON.stringify(submitSubmissionResponse));
    if (submitSubmissionResponse.isSuccess) {
      if (
        submitSubmissionResponse.responseData.submissionId != null &&
        submitSubmissionResponse.responseData.submissionId.length > 0
      ) {
        submissionId = submitSubmissionResponse.responseData.submissionId;
      } else if (
        submitSubmissionResponse.responseData.ongoingSubmissionId != null &&
        submitSubmissionResponse.responseData.ongoingSubmissionId.length > 0
      ) {
        submissionId =
          submitSubmissionResponse.responseData.ongoingSubmissionId;
      }
    }

    return new Promise<string>((resolve, reject) => {
      if (submissionId == null) {
        console.log("Failed to get submission ID");
        reject("Failed to get submission ID");
      } else {
        resolve(submissionId);
      }
    });
  }

  private LoadState() {
    this.productId = process.env[`${EnvVariablePrefix}product_id`]!;
    this.sellerId = process.env[`${EnvVariablePrefix}seller_id`]!;
    this.tenantId = process.env[`${EnvVariablePrefix}tenant_id`]!;
    this.clientId = process.env[`${EnvVariablePrefix}client_id`]!;
    this.clientSecret = process.env[`${EnvVariablePrefix}client_secret`]!;
    this.accessToken = process.env[`${EnvVariablePrefix}access_token`]!;

    console.log(`productId = ${this.productId}`);
    console.log(`sellerId = ${this.sellerId}`);
    console.log(`tenantId = ${this.tenantId}`);
    console.log(`clientId = ${this.clientId}`);
    console.log(`clientSecret = ${this.clientSecret}`);
    console.log(`accessToken = ${this.accessToken}`);
  }
}
