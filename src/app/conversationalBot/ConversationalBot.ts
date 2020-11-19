
import { BotDeclaration, MessageExtensionDeclaration, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import {
    StatePropertyAccessor,
    CardFactory, TurnContext, MemoryStorage, ConversationState,
    ActivityTypes, TeamsActivityHandler, MessageFactory,
    TaskModuleTaskInfo, TaskModuleRequest, TaskModuleResponse
} from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";

import * as Util from "util";
const TextEncoder = Util.TextEncoder;

// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for ConversationalBot Bot
 */
@BotDeclaration(
    "/api/messages2",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID_2,
    process.env.MICROSOFT_APP_PASSWORD_2)

export class ConversationalBot extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();

        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        // Set up the Activity processing

        this.onMessage(async (context: TurnContext): Promise<void> => {
            // TODO: add your own bot logic in here
            switch (context.activity.type) {
                case ActivityTypes.Message:
                    let text = TurnContext.removeRecipientMention(context.activity);
                    text = text.toLowerCase();
                    if (text.startsWith("mentionme")) {
                        if (context.activity.conversation.conversationType === "personal") {
                            await this.handleMessageMentionMeOneOnOne(context);
                        } else {
                            await this.handleMessageMentionMeChannelConversation(context);
                        }
                        return;
                    } else if (text.startsWith("hello")) {
                        await context.sendActivity("Oh, hello to you as well!");
                        return;
                    } else if (text.startsWith("help")) {
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog("help");
                    } else if (text.startsWith("exercise")) {
                        await context.sendActivity(`Exercise - Creating conversational bots for Microsoft Teams <a href="https://docs.microsoft.com/en-us/learn/modules/msteams-conversation-bots/3-exercise-conversation-bots">Link</a>`);
                    } else if (text.startsWith("learn")) {
                        const card = CardFactory.heroCard("Learn Microsoft Teams", undefined, [
                            {
                                type: "invoke",
                                title: "Watch 'Task-oriented interactions in Microsoft Teams with messaging extensions'",
                                value: { type: "task/fetch", taskModule: "player", videoId: "aHoRK8cr6Og" }
                            },
                            {
                                type: "invoke",
                                title: "Watch 'Microsoft Teams embedded web experiences'",
                                value: { type: "task/fetch", taskModule: "player", videoId: "AQcdZYkFPCY" }
                            },
                            {
                                type: "invoke",
                                title: "Watch a invalid action...",
                                value: { type: "task/fetch", taskModule: "something", videoId: "hello-world" }
                            },
                            {
                                type: "invoke",
                                title: "Watch Specific Video",
                                value: { type: "task/fetch", taskModule: "selector", videoId: "QHPBw7F4OL4" }
                            }
                        ]);
                        await context.sendActivity({ attachments: [card] });
                    } else {
                        await context.sendActivity(`I\'m terribly sorry, but my master hasn\'t trained me to do anything yet...`);
                    }
                    break;
                default:
                    break;
            }
            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });
    }
    protected handleTeamsTaskModuleFetch(context: TurnContext, request: TaskModuleRequest): Promise<TaskModuleResponse> {
        let response: TaskModuleResponse;

        switch (request.data.taskModule) {
            case "player":
                response = ({
                    task: {
                        type: "continue",
                        value: {
                            title: "YouTube Player",
                            url: `https://${process.env.HOSTNAME}/youTubePlayerTab/player.html?vid=${request.data.videoId}`,
                            width: 1000,
                            height: 700
                        } as TaskModuleTaskInfo
                    }
                } as TaskModuleResponse);
                break;
            case "selector":
                response = ({
                    task: {
                        type: "continue",
                        value: {
                            title: "YouTube Video Selector",
                            card: this.getSelectorAdaptiveCard(request.data.videoId),
                            width: 350,
                            height: 250
                        } as TaskModuleTaskInfo
                    }
                } as TaskModuleResponse);
                break;
            default:
                response = ({
                    task: {
                        type: "continue",
                        value: {
                            title: "YouTube Player",
                            url: `https://${process.env.HOSTNAME}/youTubePlayerTab/player.html?vid=X8krAMdGvCQ&default=1`,
                            width: 1000,
                            height: 700
                        } as TaskModuleTaskInfo
                    }
                } as TaskModuleResponse);
                break;
        }

        // tslint:disable-next-line: no-console
        console.log("handleTeamsTaskModuleFetch() response", response);
        return Promise.resolve(response);
    }
    protected handleTeamsTaskModuleSubmit(context: TurnContext, request: TaskModuleRequest): Promise<TaskModuleResponse> {
        const response: TaskModuleResponse = {
            task: {
                type: "continue",
                value: {
                    title: "YouTube Player",
                    url: `https://${process.env.HOSTNAME}/youTubePlayerTab/player.html?vid=${request.data.youTubeVideoId}`,
                    width: 1000,
                    height: 700
                } as TaskModuleTaskInfo
            }
        } as TaskModuleResponse;
        return Promise.resolve(response);
    }

    private getSelectorAdaptiveCard(defaultVideoId: string = "") {
        return CardFactory.adaptiveCard({
            type: "AdaptiveCard",
            version: "1.0",
            body: [
                {
                    type: "Container",
                    items: [
                        {
                            type: "TextBlock",
                            text: "YouTube Video Selector",
                            weight: "bolder",
                            size: "extraLarge"
                        }
                    ]
                },
                {
                    type: "Container",
                    items: [
                        {
                            type: "TextBlock",
                            text: "Enter the ID of a YouTube video to show in the task module player.",
                            wrap: true
                        },
                        {
                            type: "Input.Text",
                            id: "youTubeVideoId",
                            value: defaultVideoId
                        }
                    ]
                }
            ],
            actions: [
                {
                    type: "Action.Submit",
                    title: "Update"
                }
            ]
        });
    }
    private async handleMessageMentionMeOneOnOne(context: TurnContext): Promise<void> {
        const mention = {
            mentioned: context.activity.from,
            text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
            type: "mention"
        };

        const replyActivity = MessageFactory.text(`Hi ${mention.text} from a 1:1 chat.`);
        replyActivity.entities = [mention];
        await context.sendActivity(replyActivity);
    }
    private async handleMessageMentionMeChannelConversation(context: TurnContext): Promise<void> {
        const mention = {
            mentioned: context.activity.from,
            text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
            type: "mention"
        };

        const replyActivity = MessageFactory.text(`Hi ${mention.text}!`);
        replyActivity.entities = [mention];
        const followupActivity = MessageFactory.text(`*We are in a channel conversation*`);
        await context.sendActivities([replyActivity, followupActivity]);
    }

}
