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
    const userId = context.activity?.from?.aadObjectId || context.activity?.from?.id || "fallback-user";
    const wrikeToken = wrikeTokens.get(userId);

    if (!wrikeToken) {
      const loginUrl = `https://login.wrike.com/oauth2/authorize?client_id=${process.env.WRIKE_CLIENT_ID}&response_type=code&redirect_uri=${process.env.WRIKE_REDIRECT_URI}&state=${userId}`;
      return {
        task: {
          type: 'continue',
          value: {
            title: 'Login to Wrike Required',
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard',
              version: '1.5',
              body: [{ type: 'TextBlock', text: 'Please login to your Wrike account.', wrap: true }],
              actions: [{ type: 'Action.OpenUrl', title: 'Login to Wrike', url: loginUrl }]
            })
          }
        }
      };
    }

    const messageHtml = context.activity.value?.messagePayload?.body?.content || '';
    const plainTextMessage = messageHtml.replace(/<[^>]+>/g, '').trim();
    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    const descriptionField = cardJson.body.find(f => f.id === 'description');
    if (descriptionField) descriptionField.value = plainTextMessage;

    const users = await this.fetchWrikeUsers(wrikeToken);
    const folders = await this.fetchWrikeProjects(wrikeToken);

    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) userDropdown.choices = users.map(user => ({ title: user.name, value: user.id }));

    const locationDropdown = cardJson.body.find(f => f.id === 'location');
    if (locationDropdown) locationDropdown.choices = folders.map(folder => ({ title: folder.title, value: folder.id }));

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
    const userId = context.activity?.from?.aadObjectId || context.activity?.from?.id || "fallback-user";
    const wrikeToken = wrikeTokens.get(userId);

    if (!wrikeToken) {
      return {
        task: {
          type: 'message',
          value: '⚠️ Please login to Wrike first.'
        }
      };
    }

    const { title, description, assignee, location, startDate, dueDate, importance } = action.data;
    const teamsMessageLink = context.activity.value?.messagePayload?.linkToMessage || '';
    const assigneeArray = Array.isArray(assignee) ? assignee : [assignee];
    const users = await this.fetchWrikeUsers(wrikeToken);

    try {
      const response = await axios.post('https://www.wrike.com/api/v4/tasks', {
        title,
        description,
        importance,
        status: "Active",
        dates: { start: startDate, due: dueDate },
        responsibles: assigneeArray,
        parents: [location],
        customFields: [{ id: CUSTOM_FIELD_ID_TEAMS_LINK, value: teamsMessageLink }]
      }, {
        headers: { Authorization: `Bearer ${wrikeToken}` }
      });

      const task = response.data.data[0];
      const selectedUsers = users.filter(u => assigneeArray.includes(u.id));
      const assigneeNames = selectedUsers.map(u => `👤 ${u.name}`).join('\n');
      const taskLink = `https://www.wrike.com/open.htm?id=${task.id}`;
      const formattedDueDate = new Date(dueDate).toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric'
      });

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
                { type: 'TextBlock', text: '🎉 Task Created Successfully!', weight: 'Bolder', size: 'Large', color: 'Good' },
                { type: 'TextBlock', text: `**${title}**`, wrap: true },
                { type: 'TextBlock', text: 'Details:', weight: 'Bolder', spacing: 'Medium' },
                { type: 'TextBlock', text: `📅 Due Date: ${formattedDueDate}`, wrap: true },
                { type: 'TextBlock', text: `👥 Assignees:\n${assigneeNames}`, wrap: true },
                { type: 'TextBlock', text: `📊 Importance: ${importance}`, wrap: true }
              ],
              actions: [
                { type: 'Action.OpenUrl', title: '🔗 View Task in Wrike', url: taskLink }
              ]
            })
          }
        }
      };
    } catch (err) {
      console.error("❌ Wrike API Error:", err?.response?.data || err.message);
      return {
        task: {
          type: 'message',
          value: `❌ Failed to create task in Wrike.\n\n${err?.response?.data?.errorDescription || err.message}`
        }
      };
    }
  }

  async fetchWrikeUsers(token) {
    const res = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: { Authorization: `Bearer ${token}` },
      params: { deleted: false }
    });

    return res.data.data
      .filter(u => {
        const profile = u.profiles?.[0];
        return (
          profile &&
          ['User', 'Owner', 'Admin'].includes(profile.role) &&
          typeof profile.email === 'string' &&
          !profile.email.includes('wrike-robot')
        );
      })
      .map(u => ({
        id: u.id,
        name: `${u.firstName} ${u.lastName} (${u.profiles[0]?.email})`
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
    const { code, state: userId } = req.query;
    const response = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        grant_type: 'authorization_code',
        code,
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        redirect_uri: process.env.WRIKE_REDIRECT_URI
      }
    });

    wrikeTokens.set(userId, response.data.access_token);

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`
      <html>
        <head><title>Login Success</title></head>
        <body style="text-align:center;font-family:sans-serif;">
          <h2 style="color:green;">✅ Wrike login successful</h2>
          <p>You can now return to Microsoft Teams.</p>
          <a href="https://teams.microsoft.com" target="_blank"
             style="padding:10px 20px;background:#0078d4;color:white;text-decoration:none;border-radius:5px;">
            Return to Teams
          </a>
        </body>
      </html>
    `);
  } catch (err) {
    console.error('❌ OAuth Callback Error:', err?.response?.data || err.message);
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('❌ Authorization failed');
  }
});
