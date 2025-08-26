// === keep: your modules and behavior ===
const wrikeDB = require('./wrike-db'); // must exist

process.on('unhandledRejection', (reason) => {
  console.error('💥 Unhandled Promise Rejection:', reason);
});
process.on('uncaughtException', (err) => {
  console.error('💥 Uncaught Exception:', err);
});

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

// simple health check
server.get('/health', (_req, res, next) => { res.send(200, 'ok'); return next(); });

// === token cache ===
const wrikeTokens = new Map();

async function getUserToken(userId) {
  const cached = wrikeTokens.get(userId);
  if (cached) return cached;
  return new Promise((resolve) => {
    wrikeDB.loadTokens(userId, (tokens) => {
      if (tokens) wrikeTokens.set(userId, tokens);
      resolve(tokens || null);
    });
  });
}

async function refreshWrikeToken(userId) {
  const creds = await getUserToken(userId);
  if (!creds?.refreshToken) throw new Error('No refresh token available');

  const resp = await axios.post('https://login.wrike.com/oauth2/token', null, {
    params: {
      grant_type: 'refresh_token',
      refresh_token: creds.refreshToken,
      client_id: process.env.WRIKE_CLIENT_ID,
      client_secret: process.env.WRIKE_CLIENT_SECRET
    },
    timeout: 8000,
  });

  const expiresAt = Date.now() + (resp.data.expires_in * 1000);
  const next = {
    accessToken: resp.data.access_token,
    refreshToken: resp.data.refresh_token || creds.refreshToken, // sometimes absent
    expiresAt
  };
  wrikeTokens.set(userId, next);
  wrikeDB.saveTokens(userId, next.accessToken, next.refreshToken, expiresAt);
  return next.accessToken;
}

// === ensure port then listen ===
const checkPort = (port) => new Promise((resolve, reject) => {
  const s = net.createServer()
    .once('error', err => (err.code === 'EADDRINUSE' ? reject(err) : resolve()))
    .once('listening', () => s.once('close', () => resolve()).close())
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

// ================= Bot =================
class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context) {
    const userId = context.activity.from?.aadObjectId || context.activity.from?.id || 'fallback-user';
    const creds = await getUserToken(userId);
    const bufferMs = 5 * 60 * 1000;
    let token = creds?.accessToken;

    if (!token || (creds?.expiresAt && (creds.expiresAt - Date.now() < bufferMs))) {
      try {
        token = await refreshWrikeToken(userId);
      } catch {
        const loginUrl = `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${userId}`;
        return {
          task: {
            type: 'continue',
            value: {
              title: 'Login to Wrike Required',
              card: CardFactory.adaptiveCard({
                type: 'AdaptiveCard',
                version: '1.5',
                body: [{ type: 'TextBlock', text: 'Please login to Wrike.', wrap: true }],
                actions: [{ type: 'Action.OpenUrl', title: 'Login', url: loginUrl }]
              })
            }
          }
        };
      }
    }

    // robust message content extraction
    const payload = context.activity.value?.messagePayload || context.activity.messagePayload || {};
    const html = payload?.body?.content || '';
    const plain = html.replace(/<[^>]+>/g, '').trim();

    // load adaptive card
    const cardJson = JSON.parse(fs.readFileSync(path.join(__dirname, 'cards', 'taskFormCard.json'), 'utf8'));
    const descField = cardJson.body.find(f => f.id === 'description');
    if (descField) descField.value = plain;

    // Try to fetch users/projects but NEVER throw -> show login card on failure
    let users = null, folders = null;
    try {
      [users, folders] = await Promise.all([
        this.fetchWrikeUsers(token, userId),
        this.fetchWrikeProjects(token, userId),
      ]);
    } catch (e) {
      console.error('fetchTask Wrike calls failed:', e.response?.data || e.message);
    }

    if (!Array.isArray(users) || !Array.isArray(folders)) {
      const loginUrl = `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${userId}`;
      return {
        task: {
          type: 'continue',
          value: {
            title: 'Login to Wrike Required',
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard',
              version: '1.5',
              body: [{ type: 'TextBlock', text: '⚠️ Cannot reach Wrike or session expired. Please sign in.', wrap: true }],
              actions: [{ type: 'Action.OpenUrl', title: 'Login', url: loginUrl }]
            })
          }
        }
      };
    }

    // populate dropdowns
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) userDropdown.choices = users.map(u => ({ title: u.name, value: u.id }));
    const folderDropdown = cardJson.body.find(f => f.id === 'location');
    if (folderDropdown) folderDropdown.choices = folders.map(f => ({ title: f.title, value: f.id }));

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          height: 600,
          width: 600,
          card: CardFactory.adaptiveCard(cardJson)
        }
      }
    };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    const userId = context.activity.from?.aadObjectId || context.activity.from?.id || 'fallback-user';
    const creds = await getUserToken(userId);
    const bufferMs = 5 * 60 * 1000;
    let token = creds?.accessToken;

    if (!token || (creds?.expiresAt && (creds.expiresAt - Date.now() < bufferMs))) {
      try {
        token = await refreshWrikeToken(userId);
      } catch {
        const loginUrl = `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${userId}`;
        return {
          task: {
            type: 'continue',
            value: {
              title: 'Login to Wrike Required',
              card: CardFactory.adaptiveCard({
                type: 'AdaptiveCard',
                version: '1.5',
                body: [{ type: 'TextBlock', text: '⚠️ Your Wrike session expired. Please login again.', wrap: true }],
                actions: [{ type: 'Action.OpenUrl', title: 'Login', url: loginUrl }]
              })
            }
          }
        };
      }
    }

    const data = action?.data || {};
    const title = data.title || '';
    const description = data.description || '';
    const location = data.location || '';
    const startDate = data.startDate || null;
    const dueDate = data.dueDate || null;
    const importance = data.importance || 'High';

    // robust link extraction
    const link =
      context.activity.value?.messagePayload?.linkToMessage ||
      context.activity.messagePayload?.linkToMessage ||
      '';

    // fetch users for validation
    let users = await this.fetchWrikeUsers(token, userId);
    if (!Array.isArray(users)) {
      return { task: { type: 'message', value: '⚠️ Error fetching Wrike users. Please sign in again.' } };
    }

    // robust assignee parsing (avoid .includes on undefined)
    const rawAssignee = data.assignee ?? '';
    const asString = Array.isArray(rawAssignee) ? null : String(rawAssignee);
    const arr = Array.isArray(rawAssignee)
      ? rawAssignee
      : (asString.includes(',') ? asString.split(',').map(s => s.trim()).filter(Boolean) : [asString].filter(Boolean));

    const validIds = users.map(u => u.id);
    const finals = arr.filter(id => validIds.includes(id));
    if (!finals.length) {
      return { task: { type: 'message', value: '❌ Invalid assignee selected.' } };
    }

    try {
      const payload = {
        title,
        description,
        importance,
        status: 'Active',
        dates: { start: startDate, due: dueDate },
        responsibles: finals,
        parents: [location],
        customFields: CUSTOM_FIELD_ID_TEAMS_LINK && link
          ? [{ id: CUSTOM_FIELD_ID_TEAMS_LINK, value: link }]
          : []
      };

      const resp = await axios.post('https://www.wrike.com/api/v4/tasks', payload, {
        headers: { Authorization: `Bearer ${token}` },
        timeout: 10000,
      });

      const task = resp.data.data[0];
      const names = users.filter(u => finals.includes(u.id)).map(u => `👤 ${u.name}`).join('\n');
      const dueReadable = dueDate
        ? new Date(dueDate).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })
        : '—';

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
                { type: 'TextBlock', text: `📅 Due Date: ${dueReadable}`, wrap: true },
                { type: 'TextBlock', text: `👥 Assignees:\n${names}`, wrap: true },
                { type: 'TextBlock', text: `📊 Importance: ${importance}`, wrap: true }
              ],
              actions: [{ type: 'Action.OpenUrl', title: '🔗 View in Wrike', url: `https://www.wrike.com/open.htm?id=${task.id}` }]
            })
          }
        }
      };
    } catch (err) {
      return { task: { type: 'message', value: `❌ Failed to create task: ${err.response?.data?.errorDescription || err.message}` } };
    }
  }

  // ===== Wrike helpers (defensive) =====
  async fetchWrikeUsers(token, userId) {
    try {
      const r = await axios.get('https://www.wrike.com/api/v4/contacts', {
        headers: { Authorization: `Bearer ${token}` },
        params: { deleted: false },
        timeout: 8000,
      });

      // Normalize first, then filter — avoids ".includes" on undefined
      return r.data.data
        .map(u => {
          const profile = Array.isArray(u.profiles) ? u.profiles[0] : undefined;
          const email = ((profile && profile.email) || u.email || '').toString();
          const role = (profile && profile.role) ? String(profile.role) : '';
          const name = `${u.firstName || ''} ${u.lastName || ''}`.trim();
          return { id: u.id, email, role, name };
        })
        .filter(u =>
          u.email.length > 0 &&
          ['User', 'Owner', 'Admin'].includes(u.role) &&
          !u.email.toLowerCase().includes('wrike-robot')
        )
        .map(u => ({ id: u.id, name: `${u.name} (${u.email})` }));
    } catch (err) {
      if (err.response?.status === 401) return null;
      console.error('fetchWrikeUsers error:', err.response?.data || err.message);
      return null;
    }
  }

  async fetchWrikeProjects(token, userId) {
    try {
      const r = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
        headers: { Authorization: `Bearer ${token}` },
        timeout: 8000,
      });
      return (r.data.data || [])
        .filter(p => !!p.project)
        .map(p => ({ id: p.id, title: p.title }));
    } catch (err) {
      if (err.response?.status === 401) return null;
      console.error('fetchWrikeProjects error:', err.response?.data || err.message);
      return null;
    }
  }
}

const bot = new WrikeBot();

server.post('/api/messages', (req, res, next) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  })
    .then(() => next())
    .catch(err => {
      console.error('💥 processActivity error:', err);
      next(err);
    });
});

// ===== OAuth callback =====
// IMPORTANT: the "code" is SINGLE-USE. If the user refreshes this page, Wrike will return invalid_grant.
server.get('/auth/callback', async (req, res) => {
  try {
    const { code, state: userId } = req.query;
    if (!code) {
      res.writeHead(400, { 'Content-Type': 'text/plain' });
      return res.end('Missing code from Wrike');
    }

    const tr = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        grant_type: 'authorization_code',
        code,
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        redirect_uri: process.env.WRIKE_REDIRECT_URI
      },
      timeout: 8000,
    });

    const expiresAt = Date.now() + (tr.data.expires_in * 1000);
    wrikeTokens.set(userId, {
      accessToken: tr.data.access_token,
      refreshToken: tr.data.refresh_token,
      expiresAt
    });
    wrikeDB.saveTokens(userId, tr.data.access_token, tr.data.refresh_token, expiresAt);

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`
      <html>
        <body style="text-align:center;font-family:sans-serif;padding:40px;">
          <h2 style="color:green;">You have successfully logged in to Wrike</h2>
          <p style="margin-top:20px;">You may now return to Microsoft Teams to continue your task.<br/>
          <strong>Do not refresh this page</strong> (the one-time code would be reused).</p>
          <a href="https://teams.microsoft.com" style="display:inline-block;margin-top:30px;padding:10px 20px;background-color:#6264A7;color:white;text-decoration:none;border-radius:5px;">
            Open Microsoft Teams
          </a>
        </body>
      </html>
    `);
  } catch (err) {
    console.error('❌ OAuth Callback Error:', err.response?.data || err.message);
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('❌ Authorization failed');
  }
});
