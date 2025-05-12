// index.js - With Per-User Wrike OAuth Authentication
require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState, TurnContext } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');
const msal = require('@azure/msal-node');

const PORT = process.env.PORT || 3978;
const CUSTOM_FIELD_ID_TEAMS_LINK = process.env.TEAMS_LINK_CUSTOM_FIELD_ID;
const server = restify.createServer();
const userTokens = new Map(); // TEMP store per-user Wrike token (use DB in prod)

server.listen(PORT, () => {
  console.log(`✅ Bot is listening on http://localhost:${PORT}`);
});

const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context) {
    const teamsUserId = context.activity.from.aadObjectId;
    const wrikeToken = userTokens.get(teamsUserId);

    if (!wrikeToken) {
      return {
        task: {
          type: 'continue',
          value: {
            title: 'Authorize Wrike',
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard',
              body: [
                { type: 'TextBlock', text: 'Please login to Wrike to create tasks.', wrap: true },
              ],
              actions: [
                {
                  type: 'Action.OpenUrl',
                  title: 'Login to Wrike',
                  url: `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${teamsUserId}`
                }
              ],
              version: '1.5'
            })
          }
        }
      };
    }

    // Continue with task form as before...
    const messageHtml = context.activity.value?.messagePayload?.body?.content || '';
    const plainTextMessage = messageHtml.replace(/<[^>]+>/g, '').trim();
    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    const descriptionField = cardJson.body.find(f => f.id === 'description');
    if (descriptionField) descriptionField.value = plainTextMessage;

    const users = await this.fetchWrikeUsers(wrikeToken);
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) {
      userDropdown.choices = users.map(user => ({ title: user.name, value: user.id }));
    }

    const folders = await this.fetchWrikeProjects(wrikeToken);
    const locationDropdown = cardJson.body.find(f => f.id === 'location');
    if (locationDropdown) {
      locationDropdown.choices = folders.map(folder => ({ title: folder.title, value: folder.id }));
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
    const teamsUserId = context.activity.from.aadObjectId;
    const wrikeToken = userTokens.get(teamsUserId);
    if (!wrikeToken) {
      return { task: { type: "message", value: "⚠️ Please authorize Wrike first." } };
    }

    const { title, description, assignee, location, startDate, dueDate, status } = action.data;
    const response = await axios.post('https://www.wrike.com/api/v4/tasks', {
      title,
      description,
      dates: { start: startDate, due: dueDate },
      responsibles: [assignee],
      parents: [location],
      customStatusId: status
    }, {
      headers: { Authorization: `Bearer ${wrikeToken}` }
    });

    const taskLink = response.data.data[0].permalink;
    return {
      task: {
        type: 'message',
        value: `✅ Task created: [${title}](${taskLink})`
      }
    };
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
    const tokenRes = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        grant_type: 'authorization_code',
        code,
        redirect_uri: process.env.WRIKE_REDIRECT_URI,
      },
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    const accessToken = tokenRes.data.access_token;
    userTokens.set(userId, accessToken);
    res.send(200, "✅ Wrike authorization complete. You can return to Teams.");
  } catch (err) {
    console.error("❌ OAuth Error:", err.response?.data || err.message);
    res.send(500, "⚠️ Authorization failed.");
  }
});