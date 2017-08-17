const restify = require('restify');
const builder = require("botbuilder");
require('dotenv').config();

const bot = new builder.UniversalBot(
    new builder.ChatConnector({
        appId: process.env.MICROSOFT_APP_ID,
        appPassword: process.env.MICROSOFT_APP_PASSWORD
    })
);

const server = restify.createServer();
server.post('/api/messages', bot.connector('*').listen());
server.post('/api/proactive', sendProactiveMessage);
server.listen(process.env.PORT, () => {
    console.log(`${server.name} listening to ${server.url}`);
});

var savedAddress = null;

function sendProactiveMessage(request, response) {
    if (savedAddress) {
        bot.beginDialog(savedAddress, '/sendMessage', { myMessage: "Hello World" });
    }
}

bot.on('conversationUpdate', (activity) => {
    //for MS Teams you need need to look for activity.sourceEvent.eventType == 'teamMemberAdded'
    if (activity.membersAdded &&
        activity.membersAdded[0].id == activity.address.bot.id) {
        bot.beginDialog(activity.address, '/start');
    }
});

bot.dialog('/', (session) => {
    session.endDialog("Hello!");
});

bot.dialog('/start', (session) => {
    savedAddress = session.message.address;
    session.endDialog("Hello!");
});

bot.dialog('/sendMessage', (session, args, next) => {
    session.send("Hey! You have a proactive messsage, here it is: ");
    session.endDialog(args.myMessage);
});