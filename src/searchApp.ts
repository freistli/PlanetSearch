import { default as axios } from "axios";
import * as querystring from "querystring";
import planets from "./data/planets.json";
import planetModule from "./adaptiveCards/planetmodule.json";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  MessagingExtensionActionResponse,
  TaskModuleResponse
} from "botbuilder";
import * as ACData from "adaptivecards-templating";
import planetCard from "./adaptiveCards/planetcard.json";
import planetSearchModule from "./adaptiveCards/planetSearchModule.json";
import { v4 as uuidv4 } from 'uuid'
export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  public async handleTeamsTaskModuleSubmit(
    context: TurnContext,
    action: any):
    Promise<TaskModuleResponse> {
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

  };



  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const searchQuery = query.parameters[0].value;

    const attachments = [];
    planets.forEach((obj) => {
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
