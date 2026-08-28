# Microsoft Store Submission

> [!WARNING]
> ## ⚠️ This action is deprecated and no longer maintained
>
> **Please migrate to [`microsoft/microsoft-store-apppublisher`](https://github.com/microsoft/microsoft-store-apppublisher) + the [Microsoft Store Developer CLI](https://github.com/microsoft/msstore-cli).**
>
> The MSStore CLI is the supported way to publish to the Microsoft Store from CI/CD. It covers everything this
> action does — for both packaged (MSIX) and unpackaged (MSI/EXE) applications — and is actively maintained.
>
> This repository is archived. Existing workflows referencing `microsoft/store-submission@v1` will keep running,
> but no further fixes, features, or security updates will be published here.
>
> **See [Migrating to the MSStore CLI](#migrating-to-the-msstore-cli) below for a step-by-step guide.**

This is a GitHub Action to update EXE and MSI apps in the Microsoft Store.

## Migrating to the MSStore CLI

Migration is a near 1:1 mapping. Replace the `microsoft/store-submission` steps with a single setup step plus
`msstore` CLI invocations.

### 1. Replace the setup step

```yml
- uses: microsoft/microsoft-store-apppublisher@v1.4
```

This puts the `msstore` CLI on the runner's `PATH`. It works on Windows, macOS, and Linux.

### 2. Map your commands

| `microsoft/store-submission` input | MSStore CLI equivalent |
| --- | --- |
| `command: configure` with `tenant-id`, `seller-id`, `client-id`, `client-secret` | `msstore reconfigure --tenantId <id> --sellerId <id> --clientId <id> --clientSecret <secret>` |
| `command: get` | `msstore submission get <product-id>` |
| `command: get` with `module-name` / `listing-language` | `msstore submission get <product-id>` (returns the full draft) |
| `command: update` with `product-update: '<json>'` | `msstore submission update <product-id> '<json>'` |
| `command: update` with `metadata-update: '<json>'` | `msstore submission updateMetadata <product-id> '<json>'` |
| `command: poll` / `polling-submission-id` | `msstore submission poll <product-id>` |
| `command: publish` | `msstore submission publish <product-id>` |
| `type: win32` or `type: packaged` | Not needed — the CLI detects the application type automatically |

The JSON accepted by `msstore submission update` is the same Partner Center submission payload you already pass to
`product-update`, so existing payloads can be carried over as-is.

### 3. Before and after

Before:

```yml
- name: Configure Store Credentials
  uses: microsoft/store-submission@v1
  with:
    command: configure
    type: win32
    seller-id: ${{ secrets.SELLER_ID }}
    product-id: ${{ secrets.PRODUCT_ID }}
    tenant-id: ${{ secrets.TENANT_ID }}
    client-id: ${{ secrets.CLIENT_ID }}
    client-secret: ${{ secrets.CLIENT_SECRET }}

- name: Update Draft Submission
  uses: microsoft/store-submission@v1
  with:
    command: update
    product-update: '{"packages":[{"packageUrl":"https://contoso.com/App.msi","languages":["en"],"architectures":["X64"],"isSilentInstall":true}]}'

- name: Publish Submission
  uses: microsoft/store-submission@v1
  with:
    command: publish
```

After:

```yml
- name: Set up MSStore CLI
  uses: microsoft/microsoft-store-apppublisher@v1.4

- name: Configure Store Credentials
  run: >
    msstore reconfigure
    --tenantId ${{ secrets.TENANT_ID }}
    --sellerId ${{ secrets.SELLER_ID }}
    --clientId ${{ secrets.CLIENT_ID }}
    --clientSecret ${{ secrets.CLIENT_SECRET }}

- name: Update Draft Submission
  run: >
    msstore submission update ${{ secrets.PRODUCT_ID }}
    '{"packages":[{"packageUrl":"https://contoso.com/App.msi","languages":["en"],"architectures":["X64"],"isSilentInstall":true}]}'

- name: Publish Submission
  run: msstore submission publish ${{ secrets.PRODUCT_ID }}
```

### Why migrate

Beyond staying on a supported tool, the MSStore CLI also unblocks things this action never supported, including
certificate-based and federated/managed-identity authentication (see
[#20](https://github.com/microsoft/store-submission/issues/20)) instead of long-lived client secrets.

### Reference implementation

[`microsoft/PowerToys`](https://github.com/microsoft/PowerToys/blob/main/.github/workflows/msstore-submissions.yml)
migrated from this action to the MSStore CLI and is a good production example to copy from.

### More information

* [MSStore CLI documentation](https://aka.ms/msstoredevcli/docs)
* [`microsoft/microsoft-store-apppublisher`](https://github.com/microsoft/microsoft-store-apppublisher) (GitHub Action and Azure DevOps extension)
* [`microsoft/msstore-cli`](https://github.com/microsoft/msstore-cli) (the CLI itself)

---

## Quick start

1. Ensure you meet the [prerequisites](#prerequisites).

2. [Install](https://aka.ms/store-submission) the extension.

3. [Obtain](#obtaining-your-credentials) and [configure](#configuring-your-credentials) your Partner Center credentials.

4. [Add tasks](#task-reference) to your release definitions.

## Prerequisites

1. You must have an Azure AD directory, and you must have [global administrator permission](https://azure.microsoft.com/documentation/articles/active-directory-assign-admin-roles/) for the directory. You can create a new Azure AD [from Partner Center](https://msdn.microsoft.com/windows/uwp/publish/manage-account-users).

2. You must [associate your Azure AD directory with your Partner Center account](https://learn.microsoft.com/en-us/windows/apps/publish/partner-center/associate-existing-azure-ad-tenant-with-partner-center-account) to obtain the credentials to allow this extension to access your account and perform actions on your behalf.

3. The app you want to publish must already exist: this extension can only publish updates to existing applications. You can [create your app in Partner Center](https://msdn.microsoft.com/windows/uwp/publish/create-your-app-by-reserving-a-name).

4. You must have already [created at least one submission](https://msdn.microsoft.com/windows/uwp/publish/app-submissions) for your app before you can use the Publish task provided by this extension. If you have not created a submission, the task will fail.

5. More information and extra prerequisites specific to the API can be found [here](https://msdn.microsoft.com/windows/uwp/monetize/create-and-manage-submissions-using-windows-store-services).

## Obtaining your credentials

Your credentials are comprised of three parts: the Azure **Tenant ID**, the **Client ID** and the **Client secret**.
Follow these steps to obtain them:

1. In Partner Center, go to your **Account settings**, click **Manage users**, and associate your organization's Partner Center account with your organization's Azure AD directory. For detailed instructions, see [Manage account users](https://msdn.microsoft.com/windows/uwp/publish/manage-account-users).

2. In the **Manage users** page, click **Add Azure AD applications**, add the Azure AD application that represents the app or service that you will use to access submissions for your Partner Center account, and assign it the **Manager** role. If this application already exists in your Azure AD directory, you can select it on the **Add Azure AD applications** page to add it to your Partner Center account. Otherwise, you can create a new Azure AD application on the **Add Azure AD applications** page. For more information, see [Add and manage Azure AD applications](https://msdn.microsoft.com/windows/uwp/publish/manage-account-users#add-and-manage-azure-ad-applications).

3. Return to the **Manage users** page, click the name of your Azure AD application to go to the application settings, and copy the **Tenant ID** and **Client ID** values.

4. Click **Add new key**. On the following screen, copy the **Key** value, which corresponds to the **Client secret**. You *will not* be able to access this info again after you leave this page, so make sure to not lose it. For more information, see the information about managing keys in [Add and manage Azure AD applications](https://msdn.microsoft.com/windows/uwp/publish/manage-account-users#add-and-manage-azure-ad-applications).

See more details on how to create a new Azure AD application account in your organizaiton's directory and add it to your Partner Center account [here](https://docs.microsoft.com/windows/uwp/publish/add-users-groups-and-azure-ad-applications#create-a-new-azure-ad-application-account-in-your-organizations-directory-and-add-it-to-your-partner-center-account).

## Obtaining App Metadata

### Seller ID

The Seller ID can be found by clicking the gear icon in the upper right corner of the [Partner Center](https://partner.microsoft.com) and selecting `Account Settings`. You can find the Seller ID under [Legal info](https://partner.microsoft.com/dashboard/account/v3/organization/legalinfo)

### Product ID

The Product ID can be found by navigating to the overview of your application in the [Partner Center](https://partner.microsoft.com) and copying the Partner Center ID.

## Task reference

### Microsoft Store Submission

This action allows you to publish your app on the Store by creating a submission in Partner Center.

## Sample

> [!WARNING]
> The sample below uses the deprecated `microsoft/store-submission` action.
> For new workflows, see [Migrating to the MSStore CLI](#migrating-to-the-msstore-cli).

```yml
name: CI

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  start-store-submission:
    runs-on: ubuntu-latest
    steps:
      - name: Configure Store Credentials
        uses: microsoft/store-submission@v1
        with:
          command: configure
          type: win32
          seller-id: ${{ secrets.SELLER_ID }}
          product-id: ${{ secrets.PRODUCT_ID }}
          tenant-id: ${{ secrets.TENANT_ID }}
          client-id: ${{ secrets.CLIENT_ID }}
          client-secret: ${{ secrets.CLIENT_SECRET }}

      - name: Update Draft Submission
        uses: microsoft/store-submission@v1
        with:
          command: update
          product-update: '{"packages":[{"packageUrl":"https://cdn.contoso.us/prod/5.10.1.4420/ContosoIgniteInstallerFull.msi","languages":["en"],"architectures":["X64"],"isSilentInstall":true}]}'

      - name: Publish Submission
        uses: microsoft/store-submission@v1
        with:
          command: publish
```
