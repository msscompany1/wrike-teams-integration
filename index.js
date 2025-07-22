// index.js
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const https = require('https');
const net = require('net');
const restify = require('restify');
const axios = require('axios');
const wrikeDB = require('./wrike-db');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');

const PORT = process.env.PORT || 3978;
const CUSTOM_FIELD_ID_TEAMS_LINK = process.env.TEAMS_LINK_CUSTOM_FIELD_ID;

process.on('unhandledRejection', (reason) => {
  console.error('💥 Unhandled Promise Rejection:', reason);
});
process.on('uncaughtException', (err) => {
  console.error('💥 Uncaught Exception:', err);
});

try {
  require('kill-port')(PORT, 'tcp')
    .then(() => console.log(`✅ Cleared port ${PORT} before startup`))
    .catch((err) => console.warn(`⚠️ Could not clear port ${PORT}:`, err.message));
} catch (e) {
  console.warn('⚠️ kill-port not installed. Skipping port cleanup.');
}

const httpsOptions = {
  key: fs.readFileSync('/home/ubuntu/ssl/privkey.pem'),
  cert: fs.readFileSync('/home/ubuntu/ssl/fullchain.pem')
};

const server = restify.createServer(httpsOptions);
server.use(restify.plugins.queryParser());

const wrikeTokens = new Map();

async function refreshWrikeToken(userId) {
  const creds = await getUserToken(userId);
  if (!creds?.refreshToken) throw new Error('No refresh token available');
  const resp = await axios.post('https://login.wrike.com/oauth2/token', null, {
    params: {
      grant_type: 'refresh_token',
      refresh_token: creds.refreshToken,
      client_id: process.env.WRIKE_CLIENT_ID,
      client_secret: process.env.WRIKE_CLIENT_SECRET,
    }
  });
  const expiresAt = Date.now() + resp.data.expires_in * 1000;
  wrikeTokens.set(userId, {
    accessToken: resp.data.access_token,
    refreshToken: resp.data.refresh_token,
    expiresAt
  });
  await wrikeDB.saveTokens(userId, resp.data.access_token, resp.data.refresh_token, expiresAt);
  return resp.data.access_token;
}

async function getUserToken(userId) {
  let creds = wrikeTokens.get(userId);
  if (creds) return creds;
  return new Promise((resolve) => {
    wrikeDB.loadTokens(userId, (tokens) => {
      if (tokens) {
        wrikeTokens.set(userId, tokens);
        resolve(tokens);
      } else {
        resolve(null);
      }
    });
  });
}

const checkPort = (port) => new Promise((resolve, reject) => {
  const tester = net.createServer()
    .once('error', err => (err.code === 'EADDRINUSE' ? reject(err) : resolve()))
    .once('listening', () => tester.once('close', () => resolve()).close())
    .listen(port);
});

checkPort(PORT)
  .then(() => server.listen(PORT, () => console.log(`✅ HTTPS bot running on https://wrike-bot.kashida-learning.co:${PORT}`)))
  .catch(err => {
    console.error(`❌ Port ${PORT} in use:`, err.message);
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
    const userId = context.activity.from?.aadObjectId || context.activity.from?.id || 'fallback-user';
    let creds = await getUserToken(userId);
    let token = creds?.accessToken;
    const buffer = 5 * 60 * 1000;
    if (!token || (creds.expiresAt && creds.expiresAt - Date.now() < buffer)) {
      try {
        token = await refreshWrikeToken(userId);
      } catch (e) {
        return this.promptLogin(userId, 'Please login to Wrike.');
      }
    }

    const html = context.activity.value?.messagePayload?.body?.content || '';
    const plain = html.replace(/<[^>]+>/g, '').trim();
    const cardJson = JSON.parse(fs.readFileSync(path.join(__dirname, 'cards', 'taskFormCard.json'), 'utf8'));
    const descField = cardJson.body.find(f => f.id === 'description');
    if (descField) descField.value = plain;

    const users = await this.fetchWrikeUsers(token);
    const folders = await this.fetchWrikeProjects(token);

    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) userDropdown.choices = users.map(u => ({ title: u.name, value: u.id }));
    const folderDropdown = cardJson.body.find(f => f.id === 'location');
    if (folderDropdown) folderDropdown.choices = folders.map(f => ({ title: f.title, value: f.id }));

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
    const userId = context.activity.from?.aadObjectId || context.activity.from?.id || 'fallback-user';
    let creds = await getUserToken(userId);
    let token = creds?.accessToken;
    const buffer = 5 * 60 * 1000;
    if (!token || (creds.expiresAt && creds.expiresAt - Date.now() < buffer)) {
      try {
        token = await refreshWrikeToken(userId);
      } catch (e) {
        return this.promptLogin(userId, '⚠️ Your Wrike session expired. Please login again.');
      }
    }

    const { title, description, assignee, location, startDate, dueDate, importance } = action.data;
    const link = context.activity.value?.messagePayload?.linkToMessage || '';

    const users = await this.fetchWrikeUsers(token);
    const assigneeList = Array.isArray(assignee) ? assignee : assignee.split(',').map(i => i.trim());
    const validIds = users.map(u => u.id);
    const finalAssignees = assigneeList.filter(id => validIds.includes(id));

    if (!finalAssignees.length) return { task: { type: 'message', value: '❌ Invalid assignee selected.' } };

    try {
      const resp = await axios.post('https://www.wrike.com/api/v4/tasks', {
        title, description, importance, status: 'Active',
        dates: { start: startDate, due: dueDate },
        responsibles: finalAssignees,
        parents: [location],
        customFields: [{ id: CUSTOM_FIELD_ID_TEAMS_LINK, value: link }]
      }, {
        headers: { Authorization: `Bearer ${token}` }
      });

      const task = resp.data.data[0];
      const names = users.filter(u => finalAssignees.includes(u.id)).map(u => `👤 ${u.name}`).join('\n');
      const due = new Date(dueDate).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

      return {
        task: {
          type: 'continue',
          value: {
            title: '✅ Task Created',
            height: 350,
            width: 500,
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard',
              version: '1.5',
              body: [
                { type: 'TextBlock', text: '🎉 Task Created!', weight: 'Bolder', size: 'Large', color: 'Good' },
                { type: 'TextBlock', text: `**${title}**`, wrap: true },
                { type: 'TextBlock', text: `📅 Due Date: ${due}`, wrap: true },
                { type: 'TextBlock', text: `👥 Assignees:\n${names}`, wrap: true },
                { type: 'TextBlock', text: `📊 Importance: ${importance}`, wrap: true }
              ],
              actions: [{ type: 'Action.OpenUrl', title: '🔗 View in Wrike', url: `https://www.wrike.com/open.htm?id=${task.id}` }]
            })
          }
        }
      };
    } catch (err) {
      console.error('❌ Wrike API Error:', err.response?.data || err.message);
      return { task: { type: 'message', value: `❌ Failed to create task: ${err.message}` } };
    }
  }

  promptLogin(userId, message) {
    const loginUrl = `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${userId}`;
    return {
      task: {
        type: 'continue',
        value: {
          title: 'Login to Wrike Required',
          card: CardFactory.adaptiveCard({
            type: 'AdaptiveCard',
            version: '1.5',
            body: [{ type: 'TextBlock', text: message, wrap: true }],
            actions: [{ type: 'Action.OpenUrl', title: 'Login', url: loginUrl }]
          })
        }
      }
    };
  }

  async fetchWrikeUsers(token) {
    const res = await axios.get('https://www.wrike.com/api/v4/contacts', { headers: { Authorization: `Bearer ${token}` } });
    return res.data.data
      .filter(u => u.profiles?.[0]?.role !== 'Collaborator' && !u.profiles[0]?.email.includes('wrike-robot'))
      .map(u => ({ id: u.id, name: `${u.firstName} ${u.lastName} (${u.profiles[0]?.email})` }));
  }

  async fetchWrikeProjects(token) {
    const res = await axios.get('https://www.wrike.com/api/v4/folders?project=true', { headers: { Authorization: `Bearer ${token}` } });
    return res.data.data.map(f => ({ id: f.id, title: f.title }));
  }
}

const bot = new WrikeBot();

server.post('/api/messages', (req, res, next) => {
  adapter.processActivity(req, res, async (context) => await bot.run(context))
    .then(() => next())
    .catch(err => {
      console.error('💥 processActivity error:', err);
      next(err);
    });
});

server.get('/auth/callback', async (req, res) => {
  try {
    const { code, state: userId } = req.query;
    const tr = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        grant_type: 'authorization_code',
        code,
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        redirect_uri: process.env.WRIKE_REDIRECT_URI
      }
    });

    const expiresAt = Date.now() + tr.data.expires_in * 1000;
    await wrikeDB.saveTokens(userId, tr.data.access_token, tr.data.refresh_token, expiresAt);
    wrikeTokens.set(userId, { accessToken: tr.data.access_token, refreshToken: tr.data.refresh_token, expiresAt });

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`<html><body style='text-align:center;font-family:sans-serif;padding:40px;'><h2 style='color:green;'>You have successfully logged in to Wrike</h2><p>You may now return to Microsoft Teams.</p><a href='https://teams.microsoft.com' style='display:inline-block;margin-top:30px;padding:10px 20px;background-color:#6264A7;color:white;text-decoration:none;border-radius:5px;'>Open Microsoft Teams</a></body></html>`);
  } catch (err) {
    console.error('❌ OAuth Callback Error:', err.response?.data || err.message);
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('❌ Authorization failed');
  }
});
