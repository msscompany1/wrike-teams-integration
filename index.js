require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const {
  BotFrameworkAdapter,
  MemoryStorage,
  ConversationState,
} = require('botbuilder');
const {
  TeamsActivityHandler,
  CardFactory,
} = require('botbuilder');
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
  async handleTeamsMessagingExtensionFetchTask(context) {
    const rawHtml = context.activity.messagePayload?.body?.content || '';
    const messageText = rawHtml.replace(/<[^>]+>/g, '').trim();

    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    const titleField = cardJson.body.find(f => f.id === 'title');
    if (titleField) titleField.value = messageText;

    try {
      const users = await this.fetchWrikeUsers();
      const userDropdown = cardJson.body.find(f => f.id === 'assignee');
      if (userDropdown) {
        userDropdown.choices = users.map(user => ({
          title: `${user.name}`,
          value: user.id
        }));
      }
    } catch (e) {
      console.error("🔴 Failed loading Wrike users:", e.message);
    }

    try {
      const folders = await this.fetchWrikeFolders();
      const locationDropdown = cardJson.body.find(f => f.id === 'location');
      if (locationDropdown) {
        locationDropdown.choices = folders.map(f => ({
          title: f.path,
          value: f.id
        }));
      }
    } catch (e) {
      console.error("🔴 Failed loading Wrike folders:", e.message);
    }

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          height: 550,
          width: 500,
          card: CardFactory.adaptiveCard(cardJson),
        },
      },
    };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    try {
      console.log("🔁 SubmitAction received");
      console.log("🟡 Action data:", JSON.stringify(action.data, null, 2));

      const { title, assignee, location, startDate, dueDate, status } = action.data;
      if (!title || !assignee || !location) throw new Error("Missing required fields.");

      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
      const taskPayload = {
        title,
        importance: 'High',
        status,
        responsibles: [assignee],
        parents: [location],
        dates: { due: dueDate },
      };
      if (startDate) taskPayload.dates.start = startDate;

      const wrikeResponse = await axios.post('https://www.wrike.com/api/v4/tasks', taskPayload, {
        headers: {
          Authorization: `Bearer ${wrikeToken}`,
          'Content-Type': 'application/json',
        },
      });

      const task = wrikeResponse.data.data[0];
      const previewCard = CardFactory.heroCard(
        '✅ Wrike Task Created',
        `${task.title}`,
        null,
        [{ type: 'openUrl', title: 'View Task in Wrike', value: task.permalink }]
      );

      return {
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: [previewCard],
        },
      };
    } catch (error) {
      console.error("❌ Error in submitAction:", error.response?.data || error.message);
      return {
        task: {
          type: "message",
          value: `⚠️ Failed to create task: ${error.message}`,
        },
      };
    }
  }

  async fetchWrikeUsers() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: {
        Authorization: `Bearer ${wrikeToken}`,
      },
    });
    return response.data.data.filter(u => !u.deleted).map(u => ({
      id: u.id,
      name: `${u.firstName} ${u.lastName}`.trim(),
    }));
  }

  async fetchWrikeFolders() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/folders', {
      headers: {
        Authorization: `Bearer ${wrikeToken}`,
      },
    });
    return response.data.data.map(folder => ({
      id: folder.id,
      title: folder.title,
      path: (folder.project ? '📁 ' : '') + (folder.title || folder.name),
    }));
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
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
    });

    console.log("🟢 Wrike Access Token:", response.data.access_token);
    res.send(200, "✅ Wrike authorization successful. You can close this.");
  } catch (error) {
    console.error("❌ Wrike OAuth Error:", error?.response?.data || error.message);
    res.send(500, "⚠️ Failed to authorize with Wrike.");
  }
});
