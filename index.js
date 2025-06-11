// ✅ Enhanced index.js for Wrike Teams Bot
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const restify = require('restify');
const axios = require('axios');
const https = require('https');
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

server.listen(PORT, () => {
  console.log(`✅ HTTPS bot running on https://wrike-bot.kashida-learning.co:${PORT}`);
});

const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const wrikeTokens = new Map();

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
              body: [
                { type: 'TextBlock', text: 'To continue, please login to your Wrike account.', wrap: true }
              ],
              actions: [
                {
                  type: 'Action.OpenUrl',
                  title: 'Login to Wrike',
                  url: loginUrl
                }
              ]
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
    if (userDropdown) {
      userDropdown.choices = users.map(user => ({ title: user.name, value: user.id }));
    }

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

    const { title, description, assignee, location, startDate, dueDate, status, importance, comment } = action.data;
    const teamsMessageLink = context.activity.value?.messagePayload?.linkToMessage || '';

    const assigneeArray = Array.isArray(assignee) ? assignee : [assignee];

    const response = await axios.post('https://www.wrike.com/api/v4/tasks', {
      title,
      description,
      importance,
      status: "Active",
      dates: { start: startDate, due: dueDate },
      responsibles: assigneeArray,
      parents: [location],
      customFields: [
        { id: CUSTOM_FIELD_ID_TEAMS_LINK, value: teamsMessageLink }
      ]
    }, {
      headers: { Authorization: `Bearer ${wrikeToken}` }
    });

    const task = response.data.data[0];
    const taskLink = `https://www.wrike.com/open.htm?id=${task.id}`;

    const users = await this.fetchWrikeUsers(wrikeToken);
    const selectedUsers = users.filter(u => assigneeArray.includes(u.id));
    const assigneeNames = selectedUsers.map(u => `👤 ${u.name}`);

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
              { type: 'TextBlock', text: '✅ Task Created Successfully!', weight: 'Bolder', size: 'Large', color: 'Good', wrap: true },
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
                      { type: 'TextBlock', text: `📅 **Due Date:** ${formattedDueDate}`, wrap: true, spacing: 'Small' },
                      { type: 'TextBlock', text: `📊 **Importance:** ${importance}`, wrap: true, spacing: 'Small' }
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
    try {
      const wrikeResponse = await axios.get('https://www.wrike.com/api/v4/contacts', {
        params: { deleted: false },
        headers: { Authorization: `Bearer ${token}` }
      });

      return wrikeResponse.data.data.filter(w => {
        const profile = w.profiles?.[0];
        const email = profile?.email;
        const role = profile?.role;
        return email && !email.includes('wrike-robot.com') && role !== 'Collaborator';
      }).map(w => ({
        id: w.id,
        name: `${w.firstName || ''} ${w.lastName || ''}`.trim() + ` (${w.profiles[0]?.email})`
      }));
    } catch (err) {
      console.error("❌ Error in fetchWrikeUsers:", err?.response?.data || err.message);
      return [{ id: 'fallback', name: 'Fallback User' }];
    }
  }

  async fetchWrikeProjects(token) {
    const response = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: { Authorization: `Bearer ${token}` }
    });
    return response.data.data.filter(f => f.project).map(f => ({ id: f.id, title: f.title }));
  }

  async fetchGraphUsers() { return []; }
}

const bot = new WrikeBot();

server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

server.get('/auth/callback', async (req, res) => {
  try {
    const code = req.query.code;
    const userId = req.query.state;

    if (!code || !userId) {
      if (!res.headersSent) res.send(400, 'Missing code or user ID');
      return;
    }

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

    if (!res.headersSent) {
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(`
        <html>
          <head><title>Success</title></head>
          <body style="text-align:center;font-family:sans-serif;">
            <h2 style="color:green;">✅ Wrike login successful</h2>
            <p>Get Back to teams and re create the task ...</p>
            <a href="https://teams.microsoft.com" target="_blank" style="display:inline-block;margin-top:20px;padding:10px 20px;background:#28a745;color:#fff;text-decoration:none;border-radius:5px;">Return to Teams</a>
          </body>
        </html>
      `);
      return; 
    }
  } catch (err) {
    console.error('❌ OAuth Callback Error:', err?.response?.data || err.message);
    if (!res.headersSent) {
      res.writeHead(500, { 'Content-Type': 'text/plain' });
      res.end('❌ Authorization failed');
    }
  }
});
