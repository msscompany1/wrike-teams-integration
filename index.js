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
              body: [{ type: 'TextBlock', text: 'To continue, please login to your Wrike account.', wrap: true }],
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
          value: '⚠️ You must login to Wrike before creating tasks. Please try again.'
        }
      };
    }

    const { title, description, assignee, location, startDate, dueDate, importance } = action.data;
    const teamsMessageLink = context.activity.value?.messagePayload?.linkToMessage || '';
    const assigneeArray = Array.isArray(assignee) ? assignee : [assignee];
    const users = await this.fetchWrikeUsers(wrikeToken);

    let task;
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

      task = response.data.data[0];
    } catch (err) {
      console.error("❌ Wrike API Error:", err?.response?.data || err.message);
      return {
        task: {
          type: 'message',
          value: `❌ Failed to create task in Wrike.\n\n${err?.response?.data?.errorDescription || err.message}`
        }
      };
    }

    const selectedUsers = users.filter(u => assigneeArray.includes(u.id));
    const assigneeNames = selectedUsers.map(u => `👤 ${u.name}`);
    const taskLink = `https://www.wrike.com/open.htm?id=${task.id}`;
    const formattedDueDate = new Date(dueDate).toLocaleDateString('en-US', {
      year: 'numeric', month: 'long', day: 'numeric'
    });

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Task Created',
          height: 300,
          width: 450,
          card: CardFactory.adaptiveCard({
            type: 'AdaptiveCard',
            version: '1.5',
            body: [
              { type: 'TextBlock', text: '✅ Task Created Successfully!', weight: 'Bolder', size: 'Large', color: 'Good' },
              { type: 'TextBlock', text: `**${title}**`, size: 'Medium', wrap: true },
              { type: 'TextBlock', text: '📌 Task Details', weight: 'Bolder', color: 'Accent', spacing: 'Medium' },
              {
                type: 'ColumnSet',
                columns: [
                  {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                      { type: 'TextBlock', text: `👥 **Assignees:**\n${assigneeNames.join('\n')}`, wrap: true },
                      { type: 'TextBlock', text: `📅 **Due Date:** ${formattedDueDate}`, wrap: true },
                      { type: 'TextBlock', text: `📊 **Importance:** ${importance}`, wrap: true }
                    ]
                  }
                ]
              }
            ],
            actions: [
              { type: 'Action.OpenUrl', title: '🔗 View Task in Wrike', url: taskLink }
            ]
          })
        }
      }
    };
  }

  async fetchWrikeUsers(token) {
    const res = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: { Authorization: `Bearer ${token}` },
      params: { deleted: false }
    });

    return res.data.data
      .filter(u => {
        const profile = u.profiles?.[0];
        return profile && profile.role !== 'Collaborator' && !profile.email.includes('wrike-robot');
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

server.post('/api/messages', (req, res, next) => {
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
        <head><title>Success</title></head>
        <body style="text-align:center;font-family:sans-serif;">
          <h2 style="color:green;">✅ Wrike login successful</h2>
          <p>Return to Microsoft Teams and continue.</p>
          <a href="https://teams.microsoft.com" target="_blank"
            style="display:inline-block;margin-top:20px;padding:10px 20px;background:#28a745;color:#fff;text-decoration:none;border-radius:5px;">
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
