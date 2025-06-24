// ✅ Global crash protection for production
process.on('unhandledRejection', (reason) => {
  console.error('💥 Unhandled Promise Rejection:', reason);
});
process.on('uncaughtException', (err) => {
  console.error('💥 Uncaught Exception:', err);
});

// ✅ Auto-clear port 3978
try {
  require('kill-port')(3978, 'tcp')
    .then(() => console.log('✅ Cleared port 3978 before startup'))
    .catch((err) => console.warn('⚠️ Could not clear port 3978:', err.message));
} catch (e) {
  console.warn('⚠️ kill-port not installed. Skipping port cleanup.');
}

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const https = require('https');
const net = require('net');
const restify = require('restify');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');
const PORT = process.env.PORT || 3978;
const CUSTOM_FIELD_ID_TEAMS_LINK = process.env.TEAMS_LINK_CUSTOM_FIELD_ID;

const httpsOptions = {
  key: fs.readFileSync('/home/ubuntu/ssl/privkey.pem'),
  cert: fs.readFileSync('/home/ubuntu/ssl/fullchain.pem')
};

const server = restify.createServer(httpsOptions);
server.use(restify.plugins.queryParser());

const wrikeTokens = new Map();
async function refreshWrikeToken(userId) {
  const creds = wrikeTokens.get(userId);
  if (!creds?.refreshToken) {
    throw new Error('No refresh token available');
  }
  const resp = await axios.post('https://login.wrike.com/oauth2/token', null, {
    params: {
      grant_type:    'refresh_token',
      refresh_token: creds.refreshToken,
      client_id:     process.env.WRIKE_CLIENT_ID,
      client_secret: process.env.WRIKE_CLIENT_SECRET,
    }
  });
  // update stored tokens + expiry
  const expiresAt = Date.now() + (resp.data.expires_in * 1000);
  wrikeTokens.set(userId, {
    accessToken:  resp.data.access_token,
    refreshToken: resp.data.refresh_token,
    expiresAt,
  });
  return resp.data.access_token;
}
// Ensure port is free before starting
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
    const token = wrikeTokens.get(userId);
    if (!token) {
      const loginUrl = `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${userId}`;
      return {
        task: { type: 'continue', value: { title: 'Login to Wrike Required', card: CardFactory.adaptiveCard({ type: 'AdaptiveCard', version: '1.5', body: [{ type: 'TextBlock', text: 'Please login to Wrike.', wrap: true }], actions: [{ type: 'Action.OpenUrl', title: 'Login', url: loginUrl }] }) } }
      };
    }

    // Prepare task form
    const html = context.activity.value?.messagePayload?.body?.content || '';
    const plain = html.replace(/<[^>]+>/g, '').trim();
    const cardJson = JSON.parse(fs.readFileSync(path.join(__dirname, 'cards', 'taskFormCard.json'), 'utf8'));
    const descField = cardJson.body.find(f => f.id === 'description'); if (descField) descField.value = plain;

    // Fetch users and projects
    let users, folders;
    try {
      users = await this.fetchWrikeUsers(token);
      folders = await this.fetchWrikeProjects(token);
    } catch (e) {
      console.error('❌ Fetch error:', e.message);
      return { task: { type: 'message', value: '⚠️ Unable to load Wrike data. Please re-authenticate.' } };
    }

    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) userDropdown.choices = users.map(u => ({ title: u.name, value: u.id }));
    const folderDropdown = cardJson.body.find(f => f.id === 'location');
    if (folderDropdown) folderDropdown.choices = folders.map(f => ({ title: f.title, value: f.id }));

    return { task: { type: 'continue', value: { title: 'Create Wrike Task', card: CardFactory.adaptiveCard(cardJson), height: 600, width: 600 } } };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    const userId = context.activity.from?.aadObjectId || context.activity.from?.id || 'fallback-user';
    const token = wrikeTokens.get(userId);
    if (!token) return { task: { type: 'message', value: '⚠️ Please login to Wrike.' } };

    const { title, description, assignee, location, startDate, dueDate, importance } = action.data;
    const link = context.activity.value?.messagePayload?.linkToMessage || '';

    let users;
    try { users = await this.fetchWrikeUsers(token); }
    catch (e) { console.error('❌ User fetch error:', e.message); return { task: { type: 'message', value: '⚠️ Error fetching Wrike users. Please re-login.' } }; }

    const arr = Array.isArray(assignee) ? assignee : (typeof assignee === 'string' && assignee.includes(',')) ? assignee.split(',').map(i => i.trim()) : [assignee];
    const valids = users.map(u => u.id);
    const finals = arr.filter(i => valids.includes(i));
    if (!finals.length) return { task: { type: 'message', value: '❌ Invalid assignee selected.' } };

    try {
      const resp = await axios.post('https://www.wrike.com/api/v4/tasks', { title, description, importance, status: 'Active', dates: { start: startDate, due: dueDate }, responsibles: finals, parents: [location], customFields: [{ id: CUSTOM_FIELD_ID_TEAMS_LINK, value: link }] }, { headers: { Authorization: `Bearer ${token}` } });
      const task = resp.data.data[0];
      const names = users.filter(u => finals.includes(u.id)).map(u => `👤 ${u.name}`).join('\n');
      const due = new Date(dueDate).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
      return { task: { type: 'continue', value: { title: '✅ Task Created', height: 350, width: 500, card: CardFactory.adaptiveCard({ type: 'AdaptiveCard', version: '1.5', body: [ { type: 'TextBlock', text: '🎉 Task Created!', weight: 'Bolder', size: 'Large', color: 'Good' }, { type: 'TextBlock', text: `**${title}**`, wrap: true }, { type: 'TextBlock', text: `📅 Due Date: ${due}`, wrap: true }, { type: 'TextBlock', text: `👥 Assignees:\n${names}`, wrap: true }, { type: 'TextBlock', text: `📊 Importance: ${importance}`, wrap: true } ], actions: [ { type: 'Action.OpenUrl', title: '🔗 View in Wrike', url: `https://www.wrike.com/open.htm?id=${task.id}` } ] }) } } };
    } catch (err) {
      console.error('❌ Wrike API Error:', err.response?.data || err.message);
      return { task: { type: 'message', value: `❌ Failed to create task: ${err.response?.data?.errorDescription || err.message}` } };
    }
  }

  async fetchWrikeUsers(token) {
    try {
      const res = await axios.get('https://www.wrike.com/api/v4/contacts', { headers: { Authorization: `Bearer ${token}` }, params: { deleted: false } });
      return res.data.data.filter(u => { const p = u.profiles?.[0]; return p && ['User','Owner','Admin'].includes(p.role) && typeof p.email==='string' && !p.email.includes('wrike-robot'); }).map(u => ({ id: u.id, name: `${u.firstName} ${u.lastName} (${u.profiles[0]?.email||''})` }));
    } catch (err) {
      if (err.response?.status === 401) throw new Error('Wrike token expired');
      throw err;
    }
  }

  async fetchWrikeProjects(token) {
    try {
      const res = await axios.get('https://www.wrike.com/api/v4/folders?project=true', { headers: { Authorization: `Bearer ${token}` } });
      return res.data.data.filter(p => p.project).map(p => ({ id: p.id, title: p.title }));
    } catch (err) {
      if (err.response?.status === 401) throw new Error('Wrike token expired');
      throw err;
    }
  }
}

const bot = new WrikeBot();
server.post('/api/messages',
  // callback style handler: (req, res, next)
  (req, res, next) => {
    adapter.processActivity(req, res, async (context) => {
     await bot.run(context);
   })
     .then(() => next())            // signal Restify we’re done
    .catch(err => {
      console.error('💥 processActivity error:', err);
      next(err);                   // propagate error to Restify
    });
  }
 );
server.get('/auth/callback', async (req, res) => {
  try {
    const { code, state: userId } = req.query;
    const tr = await axios.post('https://login.wrike.com/oauth2/token', null, { params: { grant_type: 'authorization_code', code, client_id: process.env.WRIKE_CLIENT_ID, client_secret: process.env.WRIKE_CLIENT_SECRET, redirect_uri: process.env.WRIKE_REDIRECT_URI } });
    wrikeTokens.set(userId, tr.data.access_token);
    res.writeHead(200, {'Content-Type':'text/html'});
    res.end(`<html><body style="text-align:center;font-family:sans-serif;"><h2 style="color:green;">✅ Wrike login successful</h2><a href="https://teams.microsoft.com">Return to Teams</a></body></html>`);
  } catch (err) {
    console.error('❌ OAuth Callback Error:', err.response?.data||err.message);
    res.writeHead(500,{'Content-Type':'text/plain'});
    res.end('❌ Authorization failed');
  }
});
