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
import { LRUCache } from 'lru-cache';

export class SearchApp extends TeamsActivityHandler {

  readonly connectionName: string = config.connectionName;
  conversationState: ConversationState;
  userState: UserState;
  readonly UserProfileProperty: string = 'userProfile';
  userProfielAccessor: any;
  readonly ConversationDataProperty: string = 'conversationData';
  conversationDataAccessor: any;

  readonly cacheOptions: any = {
    max: 500,

    // for use with tracking overall storage size
    maxSize: 5000,
    sizeCalculation: (value, key) => {
      return 1
    },

    // for use when you need to clean up something when objects
    // are evicted from the cache
    dispose: (value, key) => {

    },

    // how long to live in ms
    ttl: 1000 * 60 * 5,

    // return stale items before removing from cache?
    allowStale: false,

    updateAgeOnGet: false,
    updateAgeOnHas: false,

    // async method to use for cache.fetch(), for
    // stale-while-revalidate type of behavior
    fetchMethod: async (
      key,
      staleValue,
      { options, signal, context }
    ) => { },
  }

  readonly cache = new LRUCache(this.cacheOptions);

  readonly cacheInitFlag = "Init";
  readonly cacheRevokeFlag = "Revoke";

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

    console.log("\r\nContext in Taks Module Submit: " + JSON.stringify(context));
    console.log("\r\naction in Taks Module Submit: " + JSON.stringify(action));

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

    if (action.data.signOutUserId) {

      const userID = action.data.signOutUserId;

      const cloudAdapter = context.adapter as CloudAdapter;
      const userTokenClient = context.turnState.get(cloudAdapter.UserTokenClientKey);

      console.log("\r\nContext in Taks Module Fetch: " + JSON.stringify(context));
      console.log("\r\naction in Taks Module Fetch: " + JSON.stringify(action));

      const tokeninCache = this.cache.get(userID);

      console.log("\r\nCache user id: "+ userID);
      console.log("\r\nCache Status before Sign Out: " + JSON.stringify(tokeninCache));

      await userTokenClient.signOutUser(userID, this.connectionName, context.activity.channelId);

      this.cache.set(userID, this.cacheRevokeFlag + tokeninCache);

      console.log("\r\nCache Status after Sign Out: " + JSON.stringify(this.cache.get(userID)));

      console.log("\r\nUser signOut");

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

      console.log("\r\nUser id: "+ userId);
      console.log("\r\ntokenExchangeRequest.token: " + tokenExchangeRequest.token);

      const userTokenClient = context.turnState.get(context.adapter.UserTokenClientKey);

      tokenExchangeResponse = await userTokenClient.exchangeToken(
        userId,
        this.connectionName,
        context.activity.channelId,
        { token: tokenExchangeRequest.token });

      console.log("\r\nCache Status before Token Exchange: " + JSON.stringify(this.cache.get(userId)));
      this.cache.set(userId, tokenExchangeResponse.token);
      console.log("\r\nCache Status after Token Exchange: " + JSON.stringify(this.cache.get(userId)));

      console.log('\r\ntokenExchangeResponse: ' + JSON.stringify(tokenExchangeResponse));
    }
    catch (err) {
      console.log('\r\ntokenExchange error: ' + err);
      // Ignore Exceptions
      // If token exchange failed for any reason, tokenExchangeResponse above stays null , and hence we send back a failure invoke response to the caller.
    }
    if (!tokenExchangeResponse || !tokenExchangeResponse.token) {
      return false;
    }

    console.log('\r\nExchanged token: ' + JSON.stringify(tokenExchangeResponse));
    return true;
  }

  async onInvokeActivity(context) {
    console.log('\r\nonInvoke, ' + context.activity.name);
    console.log("\r\nUser id: "+ context.activity.from.id);
    const valueObj = context.activity.value;

    if (valueObj.authentication) {

      const authObj = valueObj.authentication;
      console.log('\r\nauthObj: ' + JSON.stringify(authObj));

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

    const userTokeninCache = this.cache.get(context.activity.from.id);

    console.log("\r\nCache user id: "+ context.activity.from.id);
    console.log("\r\nCache Status in Query: " + userTokeninCache);

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

    const { signInLink } = await userTokenClient.getSignInResource(
      this.connectionName,
      context.activity
    );

    console.log("\r\nToken Response: " + JSON.stringify(tokenResponse));
    console.log("\r\nSignIn Link: " + signInLink);

    //token is not in cache means user has not signed in yet
    if (!userTokeninCache) {

      this.cache.set(context.activity.from.id, this.cacheInitFlag);

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
    //if token in cache, always update the token based on system stored user token
    else if (tokenResponse && tokenResponse.token) {

      if (userTokeninCache.toString().startsWith(this.cacheRevokeFlag) && userTokeninCache.toString().endsWith(tokenResponse.token)) {
        console.log("\r\nToken is revoked, need to sign in again");
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
      else {
        this.cache.set(context.activity.from.id, tokenResponse.token);
        console.log("\r\nCache Status updated in Query: " + JSON.stringify(this.cache.get(context.activity.from.id)));
      }
    }
    else if (!tokenResponse || !tokenResponse.token) {
      // There is no system sotred user token, so the user has not signed in yet.
      // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions

      this.cache.set(context.activity.from.id, this.cacheInitFlag);

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

    const searchQuery = query.parameters[0].value;

    const tokeninCache = this.cache.get(context.activity.from.id);
    console.log("\r\nSignIn Token in Cache: " + tokeninCache);

    const graphClient = new SimpleGraphClient(tokeninCache);
    const profile = await graphClient.GetMyProfile();
    const photo = await graphClient.GetPhotoAsync(tokeninCache);

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
            signOutUserId: context.activity.from.id,
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
