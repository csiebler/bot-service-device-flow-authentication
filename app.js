const restify = require('restify');
const builder = require('botbuilder');
const botbuilder_azure = require("botbuilder-azure");
const deviceflow = require('./deviceflow.js');
require('dotenv').config();

const BOT_STATE_TABLE = process.env.BOT_STATE_TABLE;
const STORAGE_ACCOUNT_CONNECTION = process.env.STORAGE_ACCOUNT_CONNECTION;

var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log('%s listening to %s', server.name, server.url);
});

var connector = new builder.ChatConnector({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
  openIdMetadata: process.env.BotOpenIdMetadata
});

server.post('/api/messages', connector.listen());

var azureTableClient = new botbuilder_azure.AzureTableClient(BOT_STATE_TABLE, STORAGE_ACCOUNT_CONNECTION);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

var bot = new builder.UniversalBot(connector);
bot.set('storage', tableStorage);

var menuChoices = {
  'Sign-In': 'signin',
  'Sign-Out': 'signout',
  'Query Sign-In Status': 'querySigninStatus',
  'Update access token': 'updateAccessToken',
  'Show my information': 'showUserInformation'
};

bot.dialog('/', [
  function (session) {
    builder.Prompts.choice(session, "Hello, how can I help you?", Object.keys(menuChoices), { listStyle: builder.ListStyle.button });
  },
  function (session, results) {
    var action = menuChoices[results.response.entity];
    return session.beginDialog(action);
  },
  function (session) {
    session.replaceDialog('/');
  }
]);

bot.dialog('signin', deviceflow.signInDialog);
bot.dialog('signout', deviceflow.signOutDialog);
bot.dialog('querySigninStatus', deviceflow.queryStatusDialog);
bot.dialog('updateAccessToken', deviceflow.updateAccessTokenDialog);
bot.dialog('showUserInformation', deviceflow.showUserInformationDialog);
