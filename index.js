require('dotenv').config();
const fs = require('fs');
const path = require('path');
const restify = require('restify');
const axios = require('axios');
const https = require('https');
const net = require('net');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');

const PORT = process.env.PORT || 3978;
const CUSTOM_FIELD_ID_TEAMS_LINK = process.env.TEAMS_LINK_CUSTOM_FIELD_ID;
const WRIKE_TOKEN_DIR = path.join(__dirname, 'tokens');
if (!fs.existsSync(WRIKE_TOKEN_DIR)) fs.mkdirSync(WRIKE_TOKEN_DIR);

const httpsOptions = {
  key: fs.readFileSync('/home/ubuntu/ssl/privkey.pem'),
  cert: fs.readFileSync('/home/ubuntu/ssl/fullchain.pem')
};

function checkPort(port) {
  return new Promise((resolve, reject) => {
    const tester = net.createServer()
      .once('error', err => err.code === 'EADDRINUSE' ? reject(err) : resolve())
      .once('listening', () => tester.once('close', () => resolve()).close())
      .listen(port);
  });
}

async function getValidWrikeToken(userId) {
  const tokenPath = path.join(WRIKE_TOKEN_DIR, `${userId}.json`);
  if (!fs.existsSync(tokenPath)) return null;
  const tokenData = JSON.parse(fs.readFileSync(tokenPath));
  const expired = (Date.now() - tokenData.obtained_at) > (tokenData.expires_in - 60) * 1000;
  if (!expired) return tokenData.access_token;

  try {
    const res = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        grant_type: 'refresh_token',
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        refresh_token: tokenData.refresh_token
      }
    });
    const newData = {
      access_token: res.data.access_token,
      refresh_token: res.data.refresh_token,
      expires_in: res.data.expires_in,
      obtained_at: Date.now()
    };
    fs.writeFileSync(tokenPath, JSON.stringify(newData, null, 2));
    return newData.access_token;
  } catch (err) {
    console.error('❌ Token refresh failed:', err.response?.data || err);
    return null;
  }
}

checkPort(PORT).then(() => {
  const server = restify.createServer(httpsOptions);
  server.use(restify.plugins.queryParser());

  const adapter = new BotFrameworkAdapter({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
  });

  const conversationState = new ConversationState(new MemoryStorage());

  class WrikeBot extends TeamsActivityHandler {
    constructor() {
      super();
      this.conversationState = conversationState;
    }

    async handleTeamsMessagingExtensionFetchTask(context) {
      const userId = context.activity.from.aadObjectId;
      const token = await getValidWrikeToken(userId);

      if (!token) {
        const loginUrl = `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(process.env.WRIKE_REDIRECT_URI)}&state=${userId}`;
        return {
          task: {
            type: 'continue',
            value: {
              title: 'Login Required',
              card: CardFactory.adaptiveCard({
                type: 'AdaptiveCard',
                version: '1.5',
                body: [{ type: 'TextBlock', text: 'Please login to Wrike first.' }],
                actions: [{ type: 'Action.OpenUrl', title: 'Login', url: loginUrl }]
              })
            }
          }
        };
      }

      const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
      const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

      const html = context.activity.value?.messagePayload?.body?.content || '';
      const plainText = html.replace(/<[^>]+>/g, '').trim();
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
      const userId = context.activity.from.aadObjectId;
      const token = await getValidWrikeToken(userId);
      if (!token) return { task: { type: 'message', value: '⚠️ Token expired. Please login again.' } };

      const { title, description, assignee, location, startDate, dueDate, importance } = action.data;
      try {
        const response = await axios.post(
          'https://www.wrike.com/api/v4/tasks',
          {
            title,
            description,
            importance,
            status: 'Active',
            dates: { start: startDate, due: dueDate },
            responsibles: Array.isArray(assignee) ? assignee : [assignee],
            parents: [location],
            customFields: [{ id: CUSTOM_FIELD_ID_TEAMS_LINK, value: context.activity.value.messagePayload.linkToMessage }]
          },
          { headers: { Authorization: `Bearer ${token}` } }
        );
        const task = response.data.data[0];
        return { task: { type: 'message', value: `✅ Wrike task created: ${task.title}` } };
      } catch (err) {
        console.error('❌ Task creation error:', err.response?.data || err.message);
        return { task: { type: 'message', value: '❌ Failed to create task.' } };
      }
    }

    async fetchWrikeUsers(token) {
      const { data } = await axios.get('https://www.wrike.com/api/v4/contacts', {
        headers: { Authorization: `Bearer ${token}` },
        params: { deleted: false }
      });
      return data.data.map(u => ({ id: u.id, name: `${u.firstName} ${u.lastName}`.trim() }));
    }

    async fetchWrikeProjects(token) {
      const { data } = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
        headers: { Authorization: `Bearer ${token}` }
      });
      return data.data.filter(f => f.project).map(p => ({ id: p.id, title: p.title }));
    }
  }

  const bot = new WrikeBot();

  server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async context => {
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

      const tokenData = {
        access_token: response.data.access_token,
        refresh_token: response.data.refresh_token,
        expires_in: response.data.expires_in,
        obtained_at: Date.now()
      };
      fs.writeFileSync(path.join(WRIKE_TOKEN_DIR, `${userId}.json`), JSON.stringify(tokenData, null, 2));
      res.send(200, '✅ Login successful. You can return to Teams.');
    } catch (err) {
      console.error('❌ OAuth callback failed:', err.response?.data || err.message);
      res.send(500, '❌ Login error.');
    }
  });

  server.listen(PORT, () => {
    console.log(`✅ HTTPS bot running on https://wrike-bot.kashida-learning.co:${PORT}`);
  });

}).catch(err => {
  console.error(`❌ Port ${PORT} already in use.`);
  process.exit(1);
});
