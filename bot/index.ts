// Import required packages
import * as restify from "restify";

import * as bodyParser from "body-parser";
import * as Controller from "./HomeController";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { BotFrameworkAdapter, TurnContext } from "botbuilder";

// This bot's main dialog.
import { TeamsBot } from "./teamsBot";

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
  appId: process.env.BOT_ID,
  appPassword: process.env.BOT_PASSWORD,
});

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

// Set the onTurnError for the singleton BotFrameworkAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the bot that will handle incoming messages.

var corsMiddleware = require('restify-cors-middleware');
var cors = corsMiddleware({
  preflightMaxAge: 5,
  origins: ['*'],
  allowHeaders: ['*'],
  allowMethods: ['*'],
  exposeHeaders: ['*']
});


// Create HTTP server.
const server = restify.createServer();
// server.use(restify.CORS());

// server.opts(/.*/, function (req,res,next) {
//     res.header("Access-Control-Allow-Origin", "*");
//     res.header("Access-Control-Allow-Methods", req.header("Access-Control-Request-Method"));
//     res.header("Access-Control-Allow-Headers", req.header("Access-Control-Request-Headers"));
//     res.send(200);
//     return next();
// });

server.pre(cors.preflight);
server.use(cors.actual);



server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});
server.use(bodyParser.json()); // for parsing application/json
server.use(bodyParser.urlencoded({ extended: true }));


server.post("/api/sendAgenda", Controller.syncAppData);
// Listen for incoming requests.

