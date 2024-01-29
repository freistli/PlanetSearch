# Copilot Plugin + SSO + TypeScript

This Copilot Plugin POC TypeScript project is based on traditional search-based Teams Message Extension.

I added SSO Auth features [SignIn, SignOut] after referecing the [previous TME Javascript sample](https://github.com/OfficeDev/Microsoft-Teams-Samples/tree/main/samples/msgext-search-sso-config/nodejs) and [new Copilot Plugin SSO Csharp sample](https://github.com/OfficeDev/Copilot-for-M365-Plugins-Samples/tree/7ed9f3f0dccbc86d8607da57a7a046510b0f2887/samples/msgext-product-support-sso-csharp).

The project shows Copilot Plugin can sign in as Microsoft Graph User and retrieve Graph information.

To use this projec, it requires Microsoft Entra ID app and Bot framework OAuth setup knowledge.

Comparing to non-SSO Copilot plugin, please pay attention to below differences:

1. Make sure [BOT SSO Setup](https://github.com/OfficeDev/Microsoft-Teams-Samples/blob/main/samples/bot-conversation-sso-quickstart/BotSSOSetup.md) is complated.

2. In manifest.json, it requires webApplicationInfo property (for example, if your graph SSO app is the same bot app, and this bot app exposed API in api://${{BOT_ID}} format.)

```
"webApplicationInfo": {
        "id": "${{BOT_ID}}",
        "resource": "api://${{BOT_ID}}"
      }
```
   Note: Change the settings according to your real Graph SSO app setup.

3. In manifest.json, add "token.botframework.com" to "validateDomains"

4. By default, local debug bot service doesn't include OAuth connection setting, you need to publish to Azure and then configure the Bot service.

5. Don't forget to add "CONNECTION_NAME" varabiel to the bot web application on Azure. Its value is your Bot Service OAuth connection name, which is configured at step 1.

# Test

1. Provision, deploy the project to Azure.
2. Side load the app package to Teams.
3. In Copilot Plugin setting, enable SSO Planet Dev plugin:

   <img src="https://github.com/freistli/PlanetSearch/blob/SSOAuthWithMemoryCache/Images/image.png" width="200"></img>

4. Ask the question: "ask for SSO-PlanetSearch-Dev about Jupiter"

5. Click Sign in, make sure input the user account in the tenant of your OAuth Graph App.

   <img src="https://github.com/freistli/PlanetSearch/blob/SSOAuthWithMemoryCache/Images/image1.png" width="400"></img>

6. After sign in successfully, the response card is updated to Signed in.

   <img src="https://github.com/freistli/PlanetSearch/blob/SSOAuthWithMemoryCache/Images/image2.png" width="400"></img>

7. Ask the question again:  "ask for SSO-PlanetSearch-Dev about Jupiter". We can see the resonsed card contains Jupiter info and current user info (name, email, graph photo)

   <img src="https://github.com/freistli/PlanetSearch/blob/SSOAuthWithMemoryCache/Images/image3.png" width="400"></img>

8. Click Sign Out, get Task Module response.

   <img src="https://github.com/freistli/PlanetSearch/blob/SSOAuthWithMemoryCache/Images/image4.png" width="300"></img>

9. Ask a different question, you will be prompted to sign again.

# Overview of Custom Search Results template

This app template is a search-based [message extension](https://docs.microsoft.com/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions?tabs=nodejs) that allows users to search an external system and share results through the compose message area of the Microsoft Teams client. You can now build and run your search-based message extensions in Teams, Outlook for Windows desktop and web experiences.

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 16, 18
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
> - [Set up your dev environment for extending Teams apps across Microsoft 365](https://aka.ms/teamsfx-m365-apps-prerequisites)
> Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect.
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teamsfx-cli)

1. First, select the Teams Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug (Edge)` or `Debug (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
5. To trigger the Message Extension, you can:
   1. In Teams: `@mention` Your message extension from the `search box area`, `@mention` your message extension from the `compose message area` or click the `...` under compose message area to find your message extension.
   2. In Outlook: click the `More apps` icon under compose email area to find your message extension.

**Congratulations**! You are running an application that can now search npm registries in Teams and Outlook.

![Search app demo](https://user-images.githubusercontent.com/11220663/167868361-40ffaaa3-0300-4313-ae22-0f0bab49c329.png)

## What's included in the template

| Folder       | Contents                                            |
| - | - |
| `.vscode/`    | VSCode files for debugging                          |
| `appPackage/` | Templates for the Teams application manifest        |
| `env/`        | Environment files                                   |
| `infra/`      | Templates for provisioning Azure resources          |
| `src/` | The source code for the search application |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                 | Contents                                           |
| - | - |
|`src/searchApp.ts`| Handles the business logic for this app template to query npm registry and return result list.|
|`src/index.ts`| `index.ts` is used to setup and configure the Message Extension.|

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`teamsapp.yml`|This is the main Teams Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`teamsapp.local.yml`|This overrides `teamsapp.yml` with actions that enable local execution and debugging.|

## Extend the template

Following documentation will help you to extend the template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Add single sign on to your app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-single-sign-on)
- [Access data in Microsoft Graph](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk#microsoft-graph-scenarios)
- [Use an existing Azure Active Directory application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the Teams app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft Teams app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Develop with Teams Toolkit CLI](https://aka.ms/teamsfx-cli/debug)
- [Preview the app on mobile clients](https://github.com/OfficeDev/TeamsFx/wiki/Run-and-debug-your-Teams-application-on-iOS-or-Android-client)
