require('dotenv').config();
const restify = require('restify');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');
const fs = require('fs');
const path = require('path');

// ✅ Dynamic port for Railway / Local
const PORT = process.env.PORT || 3978;

// ✅ Create HTTP server
const server = restify.createServer();
server.listen(PORT, () => {
  console.log(`✅ Bot is listening on http://localhost:${PORT}`);
});

// ✅ Create Bot Framework Adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

// ✅ Set up memory storage + state
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

// ✅ Bot logic (Wrike Teams Bot)
class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context, action) {
    const messageText = context.activity.messagePayload?.body?.content || '';

    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    // Prefill title field from selected Teams message
    const titleField = cardJson.body.find(f => f.id === 'title');
    if (titleField) titleField.value = messageText;

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          height: 400,
          width: 500,
          card: CardFactory.adaptiveCard(cardJson),
        },
      },
    };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    try {
      console.log("🔁 SubmitAction received");
      console.log("🟡 Action data:", JSON.stringify(action, null, 2));
  
      const { title, dueDate, assignee } = action.data;
  
      if (!title) {
        throw new Error("Missing required field: title");
      }
  
      const responseText = `✅ Task created: ${title} (Due: ${dueDate}) Assigned to: ${assignee}`;
      await context.sendActivity(responseText);
  
      return {
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: [],
        },
      };
    } catch (error) {
      console.error("❌ Error in submitAction:", error);
      await context.sendActivity("⚠️ Failed to create task. Try again later.");
      throw error; // This is what causes Teams to show the red "Unable to reach app"
    }
  }
}

const bot = new WrikeBot();

// ✅ Endpoint Teams uses to talk to your bot
server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// ✅ Optional test route
server.get('/', (req, res, next) => {
  res.send(200, '✔️ Railway bot is running!');
  return next();
});
