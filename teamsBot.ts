import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  MessageFactory,
} from "botbuilder";

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      // await context.sendActivity(`Echo: ${txt}`);

    //   const card = {
    //     "type": "AdaptiveCard",
    //     "contentType": "application/vnd.microsoft.card.adaptive",
    //     "body": [
    //         {
    //             "type": "TextBlock",
    //             "text": "Hello!",
    //             "size": "large"
    //         },
    //         {
    //             "type": "TextBlock",
    //             "text": "How can I help you today?",
    //             "wrap": true
    //         }
    //     ],
    //     "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    //     "version": "1.0"
    // };
    const card = {
      "type": "Action.Submit",
      "contentType": "application/vnd.microsoft.card.adaptive",
      "title": "Sign In",
      "data": {
        "msteams": {
          "type": "signin",
          "value": "https://www.bing.com/"
        },
      },
    };

    //const message = MessageFactory.attachment(card);
    //await context.sendActivity(message);

    const userCard = CardFactory.adaptiveCard(this.adaptiveCardActions());
    await context.sendActivity({ attachments: [userCard] });

    // await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            `Hi there! I'm a Teams bot that will echo what you said to me.`
          );
          break;
        }
      }
      await next();
    });
  }

  adaptiveCardActions = () => ({
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [
        {
            "type": "TextBlock",
            "text": "Adaptive Card Actions"
        }
    ],
    "actions": [
        {
            "type": "Action.OpenUrl",
            "title": "Action Open URL",
            "url": "https://adaptivecards.io"
        },
        {
          "type": "Action.Submit",
          "title": "Sign In",
          "data": {
            "msteams": {
              "type": "signin",
              "value": "https://m365tab962ca2.z5.web.core.windows.net/index.html"
              }
          }
        },
        {
            "type": "Action.ShowCard",
            "title": "Action Submit",
            "card": {
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [
                    {
                        "type": "Input.Text",
                        "id": "name",
                        "label": "Please enter your name:",
                        "isRequired": true,
                        "errorMessage": "Name is required"
                    }
                ],
                "actions": [
                    {
                        "type": "Action.Submit",
                        "title": "Submit"
                    }
                ]
            }
        },
        {
            "type": "Action.ShowCard",
            "title": "Action ShowCard",
            "card": {
                "type": "AdaptiveCard",
                "version": "1.0",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": "This card's action will show another card"
                    }
                ],
                "actions": [
                    {
                        "type": "Action.ShowCard",
                        "title": "Action.ShowCard",
                        "card": {
                            "type": "AdaptiveCard",
                            "body": [
                                {
                                    "type": "TextBlock",
                                    "text": "**Welcome To New Card**"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": "This is your new card inside another card"
                                }
                            ]
                        }
                    }
                ]
            }
        }
    ]
});
}
