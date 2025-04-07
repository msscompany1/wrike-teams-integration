// index.js - Main bot entry point
require('dotenv').config();
const restify = require('restify');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory, TurnContext } = require('botbuilder');
const fs = require('fs');
const path = require('path');

// Create server
const server = restify.createServer();
server.listen(3978, () => {
  console.log(`\nBot is listening on http://localhost:3978`);
});

// Create adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

adapter.use(async (turnContext, next) => {
  turnContext.turnState.set('conversationState', conversationState);
  await next();
});

class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context, action) {
    const messageText = context.activity.messagePayload.body.content || '';
    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf-8'));

    // Prefill the card input field with the Teams message
    const titleField = cardJson.body.find(f => f.id === 'title');
    if (titleField) titleField.value = messageText;

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          height: 300,
          width: 400,
          card: CardFactory.adaptiveCard(cardJson),
        },
      },
    };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    const { title, dueDate, assignee } = action.data;

    // TODO: Call Wrike API here to create the task
    const responseText = `✅ Wrike task created! Title: ${title}, Due: ${dueDate}, Assigned to: ${assignee}`;

    await context.sendActivity(responseText);
    return { composeExtension: { type: 'result', attachmentLayout: 'list', attachments: [] } };
  }
}

const bot = new WrikeBot();

server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
      await bot.run(context);
  });
});
server.get('/', (req, res, next) => {
  res.send(200, '✔️ Bot is running.');
  next();
});
