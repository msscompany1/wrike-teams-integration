// ✅ Load environment variables
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
    const rawMessage = context.activity.messagePayload?.body?.content || '';
    const description = rawMessage.replace(/<[^>]+>/g, '').trim();

    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    const descField = cardJson.body.find(f => f.id === 'description');
    if (descField) descField.value = description;

    const users = await this.fetchWrikeUsers();
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) {
      userDropdown.choices = users.map(user => ({
        title: user.name,
        value: user.id,
      }));
    }

    const folders = await this.fetchWrikeProjects();
    const locationDropdown = cardJson.body.find(f => f.id === 'location');
    if (locationDropdown) {
      locationDropdown.choices = folders.map(folder => ({
        title: folder.title,
        value: folder.id,
      }));
    }

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          height: 520,
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

      const { title, description, assignee, location, startDate, dueDate, status } = action.data;
      if (!title || !description) throw new Error("Missing title or description");

      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
      const payload = {
        title,
        description,
        importance: 'High',
        status,
      };

      if (startDate || dueDate) {
        payload.dates = {};
        if (startDate) payload.dates.start = startDate;
        if (dueDate) payload.dates.due = dueDate;
      }

      if (assignee) payload.responsibles = [assignee];
      if (location) payload.parents = [location];

      const wrikeResponse = await axios.post('https://www.wrike.com/api/v4/tasks', payload, {
        headers: {
          Authorization: `Bearer ${wrikeToken}`,
          'Content-Type': 'application/json',
        },
      });

      const task = wrikeResponse.data.data[0];
      console.log("✅ Wrike Task Created:", task.permalink);

      return {
        task: {
          type: "continue",
          value: {
            title: "Task Created",
            height: 150,
            width: 400,
            card: CardFactory.adaptiveCard({
              type: "AdaptiveCard",
              version: "1.4",
              body: [
                { type: "TextBlock", size: "Medium", weight: "Bolder", text: "✅ Task Created Successfully!" },
                { type: "TextBlock", wrap: true, text: task.title },
              ],
              actions: [
                {
                  type: "Action.OpenUrl",
                  title: "View Task in Wrike",
                  url: task.permalink
                }
              ]
            })
          }
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

  async fetchWrikeUsers() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: { Authorization: `Bearer ${wrikeToken}` },
    });
    return response.data.data
      .filter(u => !u.deleted)
      .map(u => ({ id: u.id, name: `${u.firstName} ${u.lastName}`.trim() }));
  }

  async fetchWrikeProjects() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: { Authorization: `Bearer ${wrikeToken}` },
    });
    return response.data.data.map(p => ({ id: p.id, title: p.title }));
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
