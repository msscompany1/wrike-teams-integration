require('dotenv').config();
const fs = require('fs');
const path = require('path');
const https = require('https');
const restify = require('restify');
const axios = require('axios');
const net = require('net');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');

const PORT = process.env.PORT || 3978;
const CUSTOM_FIELD_ID_TEAMS_LINK = process.env.TEAMS_LINK_CUSTOM_FIELD_ID;

const httpsOptions = {
  key: fs.readFileSync('/home/ubuntu/ssl/privkey.pem'),
  cert: fs.readFileSync('/home/ubuntu/ssl/fullchain.pem')
};

const wrikeTokens = new Map();

const server = restify.createServer(httpsOptions);
server.use(restify.plugins.queryParser());

const checkPort = (port) => new Promise((resolve, reject) => {
  const tester = net.createServer()
    .once('error', err => (err.code === 'EADDRINUSE' ? reject(err) : resolve()))
    .once('listening', () => tester.once('close', () => resolve()).close())
    .listen(port);
});

checkPort(PORT).then(() => {
  server.listen(PORT, () => {
    console.log(`✅ HTTPS bot running on https://wrike-bot.kashida-learning.co:${PORT}`);
  });
}).catch(err => {
  console.error(`❌ Port ${PORT} already in use. Exiting.`);
  process.exit(1);
});

const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD
});

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context) {
    const userId = context.activity.from?.aadObjectId || "unknown-user";
    const token = wrikeTokens.get(userId);

    if (!token) {
      const loginUrl = `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${userId}`;
      return {
        task: {
          type: 'continue',
          value: {
            title: 'Login to Wrike Required',
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard',
              version: '1.5',
              body: [{ type: 'TextBlock', text: 'Please login to Wrike to continue.' }],
              actions: [{ type: 'Action.OpenUrl', title: 'Login', url: loginUrl }]
            })
          }
        }
      };
    }

    const html = context.activity.value?.messagePayload?.body?.content || '';
    const plainText = html.replace(/<[^>]+>/g, '').trim();

    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));
    const descField = cardJson.body.find(f => f.id === 'description');
    if (descField) descField.value = plainText;

    const users = await this.fetchWrikeUsers(token);
    const folders = await this.fetchWrikeProjects(token);

    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) userDropdown.choices = users.map(u => ({ title: u.name, value: u.id }));

    const locDropdown = cardJson.body.find(f => f.id === 'location');
    if (locDropdown) locDropdown.choices = folders.map(f => ({ title: f.title, value: f.id }));

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
    const userId = context.activity.from?.aadObjectId || "unknown-user";
    const token = wrikeTokens.get(userId);
    if (!token) {
      return {
        task: {
          type: 'message',
          value: '⚠️ Please login to Wrike again.'
        }
      };
    }

    const { title, description, assignee, location, startDate, dueDate, importance } = action.data;
    const assigneeArray = Array.isArray(assignee) ? assignee : [assignee];
    const teamsMessageLink = context.activity.value?.messagePayload?.linkToMessage || '';

    try {
      const response = await axios.post('https://www.wrike.com/api/v4/tasks', {
        title,
        description,
        importance,
        status: 'Active',
        dates: { start: startDate, due: dueDate },
        responsibles: assigneeArray,
        parents: [location],
        customFields: [{ id: CUSTOM_FIELD_ID_TEAMS_LINK, value: teamsMessageLink }]
      }, {
        headers: { Authorization: `Bearer ${token}` }
      });

      const task = response.data.data[0];
      return {
        task: {
          type: 'message',
          value: `✅ Task created: ${task.title}`
        }
      };
    } catch (err) {
      console.error("❌ Wrike API error:", err?.response?.data || err.message);
      return {
        task: {
          type: 'message',
          value: `❌ Error creating task: ${err?.response?.data?.errorDescription || err.message}`
        }
      };
    }
  }

  async fetchWrikeUsers(token) {
    const res = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: { Authorization: `Bearer ${token}` },
      params: { deleted: false }
    });
    return res.data.data.map(u => ({
      id: u.id,
      name: `${u.firstName} ${u.lastName}`.trim()
    }));
  }

  async fetchWrikeProjects(token) {
    const res = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: { Authorization: `Bearer ${token}` }
    });
    return res.data.data.filter(p => p.project).map(p => ({ id: p.id, title: p.title }));
  }
}

const bot = new WrikeBot();

server.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

server.get('/auth/callback', async (req, res) => {
  try {
    const code = req.query.code;
    const userId = req.query.state;
    const response = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        grant_type: 'authorization_code',
        code,
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        redirect_uri: process.env.WRIKE_REDIRECT_URI
      }
    });

    const token = response.data.access_token;
    wrikeTokens.set(userId, token);

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`<html><body><h2 style="color:green;">✅ Wrike login successful</h2><p>Return to Teams to continue.</p></body></html>`);
  } catch (err) {
    console.error("❌ OAuth callback error:", err?.response?.data || err.message);
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('❌ Authorization failed');
  }
});
