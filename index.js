// ✅ Load env vars
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
const { TeamsActivityHandler, CardFactory } = require('botbuilder');
const msal = require('@azure/msal-node');

const PORT = process.env.PORT || 3978;
const server = restify.createServer();
server.listen(PORT, () => {
  console.log(`✅ Bot is listening on http://localhost:${PORT}`);
});

// Adapter & State
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
    const messageText = context.activity.messagePayload?.body?.content?.replace(/<[^>]+>/g, '').trim() || '';
    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    const descField = cardJson.body.find(f => f.id === 'description');
    if (descField) descField.value = messageText;

    const users = await this.fetchWrikeUsers();
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) {
      userDropdown.choices = users.map(u => ({ title: u.name, value: u.id }));
    }

    const folders = await this.fetchWrikeProjects();
    const locDropdown = cardJson.body.find(f => f.id === 'location');
    if (locDropdown) {
      locDropdown.choices = folders.map(f => ({ title: f.title, value: f.id }));
    }

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          height: 500,
          width: 500,
          card: CardFactory.adaptiveCard(cardJson),
        },
      },
    };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    try {
      const { title, description, assignee, location, startDate, dueDate, status } = action.data;
      console.log('🟡 SubmitAction data:', action.data);

      // ✅ Validate required fields
      const errors = [];
      if (!title) errors.push({ id: 'title', error: 'Title is required' });
      if (!location) errors.push({ id: 'location', error: 'Select a folder/project' });
      if (!assignee) errors.push({ id: 'assignee', error: 'Choose an assignee' });

      if (errors.length > 0) {
        return {
          task: {
            type: 'continue',
            value: {
              card: CardFactory.adaptiveCard({
                ...action.card,
                actions: [],
                body: action.card.body.map(field => {
                  const error = errors.find(e => e.id === field.id);
                  return error ? { ...field, isRequired: true, errorMessage: error.error } : field;
                }),
              }),
              title: 'Please fix the following fields',
              height: 500,
              width: 500,
            },
          },
        };
      }

      // ✅ Create task in Wrike
      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
      const wrikeResponse = await axios.post(
        'https://www.wrike.com/api/v4/tasks',
        {
          title,
          description,
          status,
          dates: {
            start: startDate,
            due: dueDate,
          },
          responsibles: [assignee],
          parents: [location],
        },
        {
          headers: {
            Authorization: `Bearer ${wrikeToken}`,
            'Content-Type': 'application/json',
          },
        }
      );

      const taskLink = wrikeResponse.data.data[0].permalink;
      return {
        task: {
          type: 'continue',
          value: {
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard',
              version: '1.4',
              body: [
                { type: 'TextBlock', text: '✅ Wrike Task Created!', weight: 'Bolder', size: 'Large' },
                { type: 'TextBlock', text: title, wrap: true },
                {
                  type: 'ActionSet',
                  actions: [
                    {
                      type: 'Action.OpenUrl',
                      title: '🔗 View Task in Wrike',
                      url: taskLink,
                    },
                  ],
                },
              ],
            }),
            title: 'Task Created',
            width: 400,
            height: 200,
          },
        },
      };
    } catch (error) {
      console.error('❌ Error in submitAction:', error.response?.data || error.message);
      return {
        task: {
          type: 'message',
          value: `❌ Failed to create task: ${error.message}`,
        },
      };
    }
  }

  async fetchWrikeUsers() {
    const response = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: { Authorization: `Bearer ${process.env.WRIKE_ACCESS_TOKEN}` },
    });
    return response.data.data
      .filter(u => !u.deleted)
      .map(u => ({ id: u.id, name: `${u.firstName} ${u.lastName}`.trim() }));
  }

  async fetchWrikeProjects() {
    const response = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: { Authorization: `Bearer ${process.env.WRIKE_ACCESS_TOKEN}` },
    });
    return response.data.data.map(f => ({ id: f.id, title: f.title }));
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
  if (!code) return res.send(400, 'Missing code from Wrike');

  try {
    const token = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        grant_type: 'authorization_code',
        code,
        redirect_uri: process.env.WRIKE_REDIRECT_URI,
      },
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
    console.log('🟢 Wrike OAuth success:', token.data);
    res.send(200, 'Wrike authorization successful!');
  } catch (err) {
    console.error('❌ OAuth Error:', err?.response?.data || err.message);
    res.send(500, 'Failed to authorize with Wrike');
  }
});
