import { default as axios } from "axios";
import * as querystring from "querystring";
import planets from "./data/planets.json";
import planetModule from "./adaptiveCards/planetmodule.json";
import config from "./config";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  MessagingExtensionActionResponse,
  TaskModuleResponse,
  CloudAdapter,
  UserState,
  ConversationState
} from "botbuilder";
import * as ACData from "adaptivecards-templating";
import planetCard from "./adaptiveCards/planetcard.json";
import planetSearchModule from "./adaptiveCards/planetSearchModule.json";
import { v4 as uuidv4 } from 'uuid';
import { env } from "process";
import { SimpleGraphClient } from "./simpleGraphClient";
import { UserProfile } from "./userProfile";
export class SearchApp extends TeamsActivityHandler {

  connectionName: string = config.connectionName;
  conversationState: ConversationState;
  userState: UserState;
  UserProfileProperty: string = 'userProfile';  
  userProfielAccessor: any;
  ConversationDataProperty: string = 'conversationData';
  conversationDataAccessor: any;


  /**
     *
     * @param {UserState} User state to persist configuration settings
     */
  constructor(conversationState: ConversationState, userState: UserState) {
    super();
    // Creates a new user property accessor.
    // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.

    this.userState = userState;
    this.userProfielAccessor = userState.createProperty(this.UserProfileProperty);
    this.conversationState = conversationState;
    this.conversationDataAccessor = this.conversationState.createProperty(this.ConversationDataProperty);


  }

  async run(context) {
    await super.run(context);

    // Save state changes
    await this.userState.saveChanges(context);
    await this.conversationState.saveChanges(context);

  }
  public async handleTeamsTaskModuleSubmit(
    context: TurnContext,
    action: any):
    Promise<TaskModuleResponse> {

      const userProfile = await this.userProfielAccessor.get(context, {});
      console.log("\r\nSubmit Read userProfile: " + JSON.stringify(userProfile));

      const conversationData = await this.conversationDataAccessor.get(context, {});
      console.log("\r\nSubmit Read conversationData: " + JSON.stringify(conversationData));

    if (action.data.planetSelector) {
      const searchQuery = action.data.planetSelector;
      const attachments = [];
      planets.forEach((obj) => {
        if (obj.id === searchQuery) {
          const template = new ACData.Template(planetModule);
          const card = template.expand({
            $root: {
              name: obj.name,
              summary: obj.summary,
              image: obj.imageLink,
              imageLink: obj.imageLink,
              id: obj.id,
              numSatellites: obj.numSatellites,
              solarOrbitYears: obj.solarOrbitYears,
              solarOrbitAvgDistanceKm: obj.solarOrbitAvgDistanceKm,
              imageAlt: obj.imageAlt,
              wikiLink: obj.wikiLink,
              entityId: uuidv4(),
              stageView: 'https://www.babylonjs.com/Demos/SPS/'
            },
          });
          const preview = CardFactory.heroCard(obj.name, obj.summary, [obj.imageLink]);
          const attachment = { ...CardFactory.adaptiveCard(card), preview };
          attachments.push(attachment);
        }
      });

      return {
        task: {
          type: 'continue',
          value: {
            width: 500,
            height: 450,
            title: 'Search other planets',
            card: attachments[0]
          }
        }
      };
    }
    else if (action.data.nextPlanet) {

      const template = new ACData.Template(planetSearchModule);
      const card = template.expand({
        $root: {
          Planets: planets
        },
      });
      const attachment = { ...CardFactory.adaptiveCard(card) };

      return {
        task: {
          type: 'continue',
          value: {
            width: 500,
            height: 450,
            title: 'Search other planets',
            card: attachment

          }
        }
      };

    }

    else {

      return {
        task: {
          type: 'message',
          value: "Thanks for using the app!"
        }
      }

    }
  }

  public async handleTeamsTaskModuleFetch(
    context: TurnContext,
    action: any):
    Promise<TaskModuleResponse> {

    if (action.data.signOut) {
      const cloudAdapter = context.adapter as CloudAdapter;
      const userTokenClient = context.turnState.get(cloudAdapter.UserTokenClientKey);
      
      const userProfile = await this.userProfielAccessor.get(context, {});
      console.log("\r\nRead userProfile: " + JSON.stringify(userProfile));

      const conversationData = await this.conversationDataAccessor.get(context, {});
      console.log("\r\nRead conversationData: " + JSON.stringify(conversationData));

      userProfile.tokenInvoked = true;
      await this.userState.saveChanges(context);
      console.log("\r\nUpdate signout userProfile: " + JSON.stringify(userProfile));

      conversationData.tokenInvoked = true;
      await this.conversationState.saveChanges(context);
      console.log("\r\nUpdate signout conversationData: " + JSON.stringify(conversationData));

      await userTokenClient.signOutUser(context.activity.from.id, this.connectionName, context.activity.channelId);      

      console.log("user signOut");

      const card = CardFactory.adaptiveCard({
        version: '1.0.0',
        type: 'AdaptiveCard',
        body: [
          {
            type: 'TextBlock',
            text: 'You have been signed out.'
          },
        ],
        actions: [
          {
            type: 'Action.Submit',
            title: 'Close',
            data: {
              key: 'close'
            },
          },
        ],
      });

      return {
        task: {
          type: 'continue',
          value: {
            card: card,
            height: 200,
            width: 400,
            title: 'Adaptive Card: Inputs'
          },
        },
      };
    }
    else {
      const template = new ACData.Template(planetSearchModule);
      const card = template.expand({
        $root: {
          Planets: planets
        },
      });

      const attachment = { ...CardFactory.adaptiveCard(card) };

      return {
        task: {
          type: 'continue',
          value: {
            width: 500,
            height: 450,
            title: 'Search other planets',
            card: attachment

          }
        }
      };
    }
  };

  async tokenIsExchangeable(context) {
    let tokenExchangeResponse = null;
    try {
      const userId = context.activity.from.id;
      const valueObj = context.activity.value;
      const tokenExchangeRequest = valueObj.authentication;
      console.log("tokenExchangeRequest.token: " + tokenExchangeRequest.token);

      const userTokenClient = context.turnState.get(context.adapter.UserTokenClientKey);

      tokenExchangeResponse = await userTokenClient.exchangeToken(
        userId,
        this.connectionName,
        context.activity.channelId,
        { token: tokenExchangeRequest.token });

      console.log('tokenExchangeResponse: ' + JSON.stringify(tokenExchangeResponse));
    }
    catch (err) {
      console.log('tokenExchange error: ' + err);
      // Ignore Exceptions
      // If token exchange failed for any reason, tokenExchangeResponse above stays null , and hence we send back a failure invoke response to the caller.
    }
    if (!tokenExchangeResponse || !tokenExchangeResponse.token) {
      return false;
    }

    console.log('Exchanged token: ' + JSON.stringify(tokenExchangeResponse));
    return true;
  }

  async onInvokeActivity(context) {
    console.log('onInvoke, ' + context.activity.name);
    const valueObj = context.activity.value;
    if (valueObj.authentication) {
      const authObj = valueObj.authentication;
      console.log('authObj: ' + JSON.stringify(authObj));
      if (authObj.token) {
        // If the token is NOT exchangeable, then do NOT deduplicate requests.
        if (await this.tokenIsExchangeable(context)) {
          return await super.onInvokeActivity(context);
        }
        else {
          const response =
          {
            status: 412
          };
          return response;
        }
      }
    }


    return await super.onInvokeActivity(context);

  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {

    console.log("\r\ncontext: " + JSON.stringify(context));
    console.log("\r\nquery: " + JSON.stringify(query));

    const userProfile = await this.userProfielAccessor.get(context, {});
    console.log("\r\nRead userProfile: " + JSON.stringify(userProfile));

    const conversationData = await this.conversationDataAccessor.get(context, {});
    console.log("\r\nRead conversationData: " + JSON.stringify(conversationData));


    const cloudAdapter = context.adapter as CloudAdapter;

    const userTokenClient = context.turnState.get(cloudAdapter.UserTokenClientKey);
    const magicCode =
      query.state && Number.isInteger(Number(query.state))
        ? query.state
        : '';
    
    
    const tokenResponse = await userTokenClient.getUserToken(
          context.activity.from.id,
          this.connectionName,
          context.activity.channelId,
          magicCode
        );

    if (userProfile.tokenInvoked && userProfile.token === tokenResponse.token) {

      
    console.log("\r\nToken revoked userState: " + JSON.stringify(this.userState));

    console.log("\r\nToken revoked conversationState: " + JSON.stringify(this.conversationState));

     
      const { signInLink } = await userTokenClient.getSignInResource(
        this.connectionName,
        context.activity
      );

    this.userState.saveChanges(context);

      return {
        composeExtension: {
          type: 'auth',
          suggestedActions: {
            actions: [
              {
                type: 'openUrl',
                value: signInLink,
                title: 'Bot Service OAuth'
              },
            ],
          },
        },
      };
    }

    userProfile.token = tokenResponse.token;
    userProfile.tokenInvoked = false;
    this.userState.saveChanges(context);

    conversationData.token = tokenResponse.token;
    conversationData.tokenInvoked = false;
    this.conversationState.saveChanges(context);

    
    console.log("\r\nToken accepted userState: " + JSON.stringify(this.userState));

    console.log("\r\nToken accepted conversationState: " + JSON.stringify(this.conversationState));


    if (!tokenResponse || !tokenResponse.token) {
      // There is no token, so the user has not signed in yet.
      // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
      const { signInLink } = await userTokenClient.getSignInResource(
        this.connectionName,
        context.activity
      );

      return {
        composeExtension: {
          type: 'auth',
          suggestedActions: {
            actions: [
              {
                type: 'openUrl',
                value: signInLink,
                title: 'Bot Service OAuth'
              },
            ],
          },
        },
      };
    }

    userProfile.tokenInvoked = false;
    userProfile.token = tokenResponse.token;
    await this.userState.saveChanges(context);
    console.log("\r\nUpdate signin userProfile: " + JSON.stringify(userProfile));

    conversationData.tokenInvoked = false;
    conversationData.token = tokenResponse.token;
    await this.conversationState.saveChanges(context);
    console.log("\r\nUpdate signin conversationData: " + JSON.stringify(conversationData));

    const searchQuery = query.parameters[0].value;

    const graphClient = new SimpleGraphClient(tokenResponse.token);
    const profile = await graphClient.GetMyProfile();
    const photo = await graphClient.GetPhotoAsync(tokenResponse.token);

    console.log("profile: " + JSON.stringify(profile));

    const attachments = [];
    planets.forEach(async (obj) => {
      if (obj.name.toLowerCase().includes(searchQuery.toLowerCase())) {
        const template = new ACData.Template(planetCard);
        const card = template.expand({
          $root: {
            name: obj.name,
            summary: obj.summary,
            image: obj.imageLink,
            imageLink: obj.imageLink,
            id: obj.id,
            numSatellites: obj.numSatellites,
            solarOrbitYears: obj.solarOrbitYears,
            solarOrbitAvgDistanceKm: obj.solarOrbitAvgDistanceKm,
            imageAlt: obj.imageAlt,
            currentVisitor: profile.displayName + " " + profile.mail,
            visitorPhoto: photo,
            wikiLink: obj.wikiLink,
            entityId: uuidv4(),
            stageView: searchQuery.toLowerCase() === 'saturn' ? 'https://www.babylonjs.com/Demos/SPS/' : 'https://www.babylonjs-playground.com/frame.html#KEKCLV'
          },
        });
        const preview = CardFactory.heroCard(obj.name, obj.summary, [obj.imageLink]);
        const attachment = { ...CardFactory.adaptiveCard(card), preview };
        attachments.push(attachment);

      }
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }
}
