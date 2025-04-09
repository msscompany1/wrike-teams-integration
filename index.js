require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');
const msal = require('@azure/msal-node');

// ✅ Setup Port
const PORT = process.env.PORT || 3978;

// ✅ Setup Restify Server
const server = restify.createServer();
server.listen(PORT, () => {
  console.log(`✅ Bot is listening on http://localhost:${PORT}`);
});

// ✅ Bot Adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

// ✅ State Storage
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

// ✅ MSAL Config
const msalConfig = {
  auth: {
    clientId: process.env.MICROSOFT_APP_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.MICROSOFT_APP_PASSWORD,
  },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

// ✅ Teams Bot Logic
class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context) {
    const messageText = context.activity.messagePayload?.body?.content || '';

    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

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
      if (!title) throw new Error("Title is required");

      // ✅ Acquire MS Graph Token
      const tokenResponse = await cca.acquireTokenByClientCredential({
        scopes: ["https://graph.microsoft.com/.default"],
      });

      console.log("🟢 MS Graph Token acquired:", !!tokenResponse.accessToken);

      // ✅ Create Wrike Task
      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;

      const response = await axios.post('https://www.wrike.com/api/v4/tasks', null, {
        headers: {
          Authorization: `Bearer ${wrikeToken}`,
        },
        params: {
          title,
          importance: 'High',
        },
      });

      const wrikeTask = response.data.data[0];
      const wrikeUrl = wrikeTask.permalink;

      console.log("🟢 Wrike Task Created:", wrikeUrl);

      await context.sendActivity(`✅ Task created in Wrike: [${title}](${wrikeUrl})`);
      return {
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: [],
        },
      };
    } catch (error) {
      console.error("❌ Error in submitAction:", error.response?.data || error.message);
      await context.sendActivity("⚠️ Failed to create task. Try again later.");
      throw error;
    }
  }
}

const bot = new WrikeBot();

// ✅ POST endpoint for Teams
server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// ✅ Basic GET endpoint
server.get('/', (req, res, next) => {
  res.send(200, '✔️ Railway bot is running!');
  return next();
});

// ✅ Wrike OAuth Callback (Fixed!)
server.get('/auth/callback', (req, res, next) => {
  const code = req.query.code;
  if (!code) {
    res.send(400, 'Missing authorization code');
    return next();
  }

  axios.post('https://login.wrike.com/oauth2/token', null, {
    params: {
      client_id: process.env.WRIKE_CLIENT_ID,
      client_secret: process.env.WRIKE_CLIENT_SECRET,
      grant_type: 'authorization_code',
      code,
      redirect_uri: process.env.WRIKE_REDIRECT_URI,
    },
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
  }).then(response => {
    console.log("🟢 Wrike Access Token:", response.data.access_token);
    res.send(200, '✅ Wrike OAuth success. You may close this window.');
    return next();
  }).catch(error => {
    console.error("❌ Wrike OAuth error:", error?.response?.data || error.message);
    res.send(500, '⚠️ Wrike OAuth failed.');
    return next();
  });
});
