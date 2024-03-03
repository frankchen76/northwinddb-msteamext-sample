// Import required packages
import * as restify from "restify";
const path = require("path");

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
    CloudAdapter,
    ConfigurationServiceClientCredentialFactory,
    ConfigurationBotFrameworkAuthentication,
    TurnContext,
    MemoryStorage,
    MessagingExtensionAttachment,
    ActivityTypes,
} from "botbuilder";
import { ApplicationBuilder, TurnState, TeamsAdapter, AuthError, Query } from '@microsoft/teams-ai';

// This bot's main dialog.
// import { SearchApp } from "./searchApp";
import config from "./config";
// import { TeamsBot } from "./sso/handleProductSearch";
import { TeamAICredential, TeamsExtService } from "./sso/TeamsExtService";
import { log } from './log';

// // Create adapter.
// // See https://aka.ms/about-bot-adapter to learn more about adapters.
// const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
//     MicrosoftAppId: config.botId,
//     MicrosoftAppPassword: config.botPassword,
//     MicrosoftAppType: "MultiTenant",
// });

// const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
//     {},
//     credentialsFactory
// );
// const adapter = new CloudAdapter(botFrameworkAuthentication);

const adapter = new TeamsAdapter(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: config.botId,
        MicrosoftAppPassword: config.botPassword,
        MicrosoftAppType: 'MultiTenant'
    })
);


// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        "OnTurnError Trace",
        `${error}`,
        "https://www.botframework.com/schemas/error",
        "TurnError"
    );

    // Send a message to the user
    await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the bot that will handle incoming messages.
// const searchApp = new SearchApp();
// const bot = new TeamsBot();

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

//log('env:', process.env);
log(`clienId: ${process.env.M365_CLIENT_ID}`);
log(`clientSecret: ${process.env.M365_CLIENT_SECRET}`);
log(`authority: ${process.env.M365_AUTHORITY_HOST}/${process.env.M365_TENANT_ID}`);
log(`signinlink: ${process.env.INITIATE_LOGIN_ENDPOINT}`);

const storage = new MemoryStorage();
const app = new ApplicationBuilder()
    .withStorage(storage)
    .withAuthentication(adapter, {
        settings: {
            graph: {
                scopes: ['api://ab4e8ed8-c4d7-4de9-a352-d23da8651cf9/.default'],
                //scopes: ['User.Read'],
                msalConfig: {
                    auth: {
                        clientId: process.env.M365_CLIENT_ID!,
                        clientSecret: process.env.M365_CLIENT_SECRET!,
                        authority: `${process.env.M365_AUTHORITY_HOST}/${process.env.M365_TENANT_ID}`
                    }
                },
                //signInLink: `https://${process.env.BOT_DOMAIN}/auth-start.html`,
                signInLink: `${process.env.INITIATE_LOGIN_ENDPOINT}`,
                endOnInvalidMessage: true
            }
        },
        autoSignIn: (context: TurnContext) => {
            log('autoSignIn', context.activity?.value?.commandId);
            const signOutActivity = context.activity?.value?.commandId === 'signOutCommand';
            if (signOutActivity) {
                return Promise.resolve(false);
            }

            return Promise.resolve(true);
        }
    })
    .build();

// Handles when the user makes a Messaging Extension query.
app.messageExtensions.query('inventorySearch', async (_context: TurnContext, state: TurnState, query: Query<Record<string, string>>) => {
    //log('env:', process.env);
    const token = state.temp.authTokens['graph'];
    const extService = new TeamsExtService();
    if (!token) {
        throw new Error('No auth token found in state. Authentication failed.');
    }
    const cred = new TeamAICredential(token);

    const ret = await extService.searchProductsFromCopilot(_context, query, cred);
    return ret.composeExtension;

});

// Listen for item selection
// app.messageExtensions.selectItem(async (_context: TurnContext, _state: TurnState, item) => {
//     // Generate detailed result
//     const card = createNpmPackageCard(item);

//     // Return results
//     return {
//         attachmentLayout: 'list',
//         attachments: [card],
//         type: 'result'
//     } as MessagingExtensionResult;
// });


// // Listen for incoming requests.
// server.post("/api/messages", async (req, res) => {
//   await adapter.process(req, res, async (context) => {
//     await searchApp.run(context);
//   });
// });

interface ConversationState {
    count: number;
}
type ApplicationTurnState = TurnState<ConversationState>;

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
app.activity(ActivityTypes.Message, async (context: TurnContext, state: ApplicationTurnState) => {
    // Increment count state
    let count = state.conversation.count ?? 0;
    state.conversation.count = ++count;

    log(state.temp.authTokens['graph']);

    // Echo back users request
    await context.sendActivity(`[${count}] you said: ${context.activity.text}`);
});

app.authentication.get('graph').onUserSignInSuccess(async (context: TurnContext, state: ApplicationTurnState) => {
    // Successfully logged in
    await context.sendActivity('Successfully logged in');
    await context.sendActivity(`Token string length: ${state.temp.authTokens['graph']!.length}`);
    await context.sendActivity(`This is what you said before the AuthFlow started: ${context.activity.text}`);
});

app.authentication
    .get('graph')
    .onUserSignInFailure(async (context: TurnContext, _state: ApplicationTurnState, error: AuthError) => {
        // Failed to login
        await context.sendActivity('Failed to login');
        await context.sendActivity(`Error message: ${error.message}`);
    });


// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
    await adapter.process(req, res, async (context) => {
        //await searchApp.run(context);
        //await bot.run(context);
        await app.run(context);
    });
    // .catch((err) => {
    //     // Error message including "412" means it is waiting for user's consent, which is a normal process of SSO, sholdn't throw this error.
    //     if (!err.message.includes("412")) {
    //         throw err;
    //     }
    // })
});

server.get(
    "/auth-:name(start|end).html",
    restify.plugins.serveStatic({
        directory: path.join(__dirname, "public"),
    })
);