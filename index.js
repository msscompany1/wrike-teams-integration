// index.js - Form-first, login-after Wrike Task Flow with conditional dropdown loading
require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState, TurnContext } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');

const PORT = process.env.PORT || 3978;
const CUSTOM_FIELD_ID_TEAMS_LINK = process.env.TEAMS_LINK_CUSTOM_FIELD_ID;
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
const wrikeTokens = new Map();
const pendingSubmissions = new Map();

class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context) {
    const userId = context.activity.from.aadObjectId;
    const wrikeToken = wrikeTokens.get(userId);
    const messageHtml = context.activity.value?.messagePayload?.body?.content || '';
    const plainTextMessage = messageHtml.replace(/<[^>]+>/g, '').trim();
    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    const descriptionField = cardJson.body.find(f => f.id === 'description');
    if (descriptionField) descriptionField.value = plainTextMessage;

    // Populate dropdowns if token exists
    if (wrikeToken) {
      const users = await this.fetchWrikeUsers(wrikeToken);
      const folders = await this.fetchWrikeProjects(wrikeToken);

      const userDropdown = cardJson.body.find(f => f.id === 'assignee');
      if (userDropdown) {
        userDropdown.choices = users.map(user => ({ title: user.name, value: user.id }));
      }

      const locationDropdown = cardJson.body.find(f => f.id === 'location');
      if (locationDropdown) {
        locationDropdown.choices = folders.map(folder => ({ title: folder.title, value: folder.id }));
      }
    }

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          card: CardFactory.adaptiveCard(cardJson),
          height: 600,
          width: 600
        }
      }
    };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    const userId = context.activity.from.aadObjectId;
    const wrikeToken = wrikeTokens.get(userId);
    const formData = action.data;

    if (!wrikeToken) {
      pendingSubmissions.set(userId, formData);

      return {
        task: {
          type: 'continue',
          value: {
            title: 'Authorize Wrike',
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard',
              version: '1.5',
              body: [
                { type: 'TextBlock', text: 'Please login to Wrike to continue task creation.', wrap: true }
              ],
              actions: [
                {
                  type: 'Action.OpenUrl',
                  title: 'Login to Wrike',
                  url: `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${userId}`
                }
              ]
            })
          }
        }
      };
    }

    return await createTask(formData, wrikeToken);
  }

  async fetchWrikeUsers(token) {
    const res = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: { Authorization: `Bearer ${token}` }
    });
    return res.data.data.map(u => ({ id: u.id, name: `${u.firstName} ${u.lastName}`.trim() }));
  }

  async fetchWrikeProjects(token) {
    const res = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: { Authorization: `Bearer ${token}` }
    });
    return res.data.data.map(f => ({ id: f.id, title: f.title }));
  }
}

async function createTask(formData, token) {
  const { title, description, assignee, location, startDate, dueDate, status } = formData;
  const response = await axios.post('https://www.wrike.com/api/v4/tasks', {
    title,
    description,
    dates: { start: startDate, due: dueDate },
    responsibles: [assignee],
    parents: [location],
    customStatusId: status
  }, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const task = response.data.data[0];
  return {
    task: {
      type: 'continue',
      value: {
        title: 'Task Created',
        height: 250,
        width: 400,
        card: CardFactory.adaptiveCard({
          type: 'AdaptiveCard',
          version: '1.5',
          body: [
            { type: 'TextBlock', text: '✅ Task created successfully!', weight: 'Bolder', size: 'Medium' },
            { type: 'TextBlock', text: `🔗 ${task.title}`, wrap: true },
            { type: 'TextBlock', text: `View: ${task.permalink}`, wrap: true }
          ]
        })
      }
    }
  };
}

const bot = new WrikeBot();

server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

server.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  const userId = req.query.state;

  if (!code || !userId) return res.send(400, 'Missing code or user ID');

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

    const token = response.data.access_token;
    wrikeTokens.set(userId, token);

    const formData = pendingSubmissions.get(userId);
    if (formData) {
      const result = await createTask(formData, token);
      pendingSubmissions.delete(userId);
      res.send(200, '✅ Wrike task created. Return to Teams to view confirmation.');
    } else {
      res.send(200, '✅ Wrike login successful. You may now return to Teams.');
    }
  } catch (err) {
    console.error('❌ OAuth Callback Error:', err.response?.data || err.message);
    res.send(500, '❌ Authorization failed');
  }
});
