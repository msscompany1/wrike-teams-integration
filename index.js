require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');
const msal = require('@azure/msal-node');

const PORT = process.env.PORT || 3978;
const server = restify.createServer();
server.listen(PORT, () => {
  console.log(`✅ Bot is listening on http://localhost:${PORT}`);
});

const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

const msalConfig = {
  auth: {
    clientId: process.env.MICROSOFT_APP_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.MICROSOFT_APP_PASSWORD,
  },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context, action) {
    const messageText = context.activity.messagePayload?.body?.content || '';
    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    // Populate Wrike spaces/folders
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const foldersRes = await axios.get('https://www.wrike.com/api/v4/folders', {
      headers: { Authorization: `Bearer ${wrikeToken}` },
    });

    const folderChoices = foldersRes.data.data.map(f => ({
      title: f.title,
      value: f.id
    }));

    const locationDropdown = cardJson.body.find(f => f.id === 'location');
    if (locationDropdown) locationDropdown.choices = folderChoices;

    // Populate assignees
    const contactsRes = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: { Authorization: `Bearer ${wrikeToken}` },
    });

    const userChoices = contactsRes.data.data
      .filter(u => !u.deleted)
      .map(u => ({
        title: `${u.firstName} ${u.lastName || ''}`.trim(),
        value: u.id
      }));

    const assigneeDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (assigneeDropdown) assigneeDropdown.choices = userChoices;

    // Set message as title default
    const titleField = cardJson.body.find(f => f.id === 'title');
    if (titleField) titleField.value = messageText.replace(/<[^>]+>/g, '');

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          height: 500,
          width: 600,
          card: CardFactory.adaptiveCard(cardJson),
        },
      },
    };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    try {
      console.log("🔁 SubmitAction received");
      console.log("🟡 Action data:", JSON.stringify(action, null, 2));

      const { title, location, assignee, startDate, dueDate, status } = action.data;
      if (!title || !location || !assignee) throw new Error("Missing required fields");

      const tokenResponse = await cca.acquireTokenByClientCredential({
        scopes: ["https://graph.microsoft.com/.default"],
      });
      console.log("🟢 MS Graph Token acquired:", tokenResponse.accessToken ? "✅" : "❌");

      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
      const taskRes = await axios.post(`https://www.wrike.com/api/v4/folders/${location}/tasks`, null, {
        headers: { Authorization: `Bearer ${wrikeToken}` },
        params: {
          title,
          status,
          importance: 'High',
          responsibles: assignee,
          dates: {
            start: startDate,
            due: dueDate
          }
        },
      });

      const permalink = taskRes.data.data[0].permalink;
      return {
        task: {
          type: "message",
          value: `✅ Task created in Wrike: [${title}](${permalink})`
        }
      };
    } catch (error) {
      console.error("❌ Error in submitAction:", error.response?.data || error.message);
      return {
        task: {
          type: "message",
          value: `⚠️ Failed to create task: ${error.message}`
        }
      };
    }
  }
}

const bot = new WrikeBot();

server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

server.get('/', (req, res, next) => {
  res.send(200, '✔️ Railway bot is running!');
  return next();
});

server.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  if (!code) {
    res.send(400, 'Missing code from Wrike');
    return;
  }

  try {
    const response = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        grant_type: 'authorization_code',
        code,
        redirect_uri: process.env.WRIKE_REDIRECT_URI,
      },
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });

    console.log("🟢 Wrike Access Token:", response.data.access_token);
    res.send(200, "✅ Wrike authorization successful. You can close this.");
  } catch (error) {
    console.error("❌ Wrike OAuth Error:", error?.response?.data || error.message);
    res.send(500, "⚠️ Failed to authorize with Wrike.");
  }
});
