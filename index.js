// Updated index.js with fixed Wrike + Graph matching
require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');
const msal = require('@azure/msal-node');

const PORT = process.env.PORT || 3978;

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

const msalConfig = {
  auth: {
    clientId: process.env.MICROSOFT_APP_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.MICROSOFT_APP_PASSWORD,
  },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context) {
    const messageHtml = context.activity.value?.messagePayload?.body?.content || '';
    const plainTextMessage = messageHtml.replace(/<[^>]+>/g, '').trim();
    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    const descriptionField = cardJson.body.find(f => f.id === 'description');
    if (descriptionField) {
      descriptionField.value = plainTextMessage;
    }
    console.log("🟢 messagePayload:", context.activity.value?.messagePayload);

    const users = await this.fetchWrikeUsers();
    console.log("✅ Wrike Users returned:", users);
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) {
      userDropdown.choices = users.map(user => ({
        title: user.name || 'Unknown',
        value: user.id,
      }));
    }

    const folders = await this.fetchWrikeProjects();
    const locationDropdown = cardJson.body.find(f => f.id === 'location');
    if (locationDropdown) {
      locationDropdown.choices = folders.map(folder => ({
        title: folder.title,
        value: folder.id,
      }));
    }

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Create Wrike Task',
          height: 600,
          width: 600,
          card: CardFactory.adaptiveCard(cardJson),
        },
      },
    };
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    try {
      console.log("🔁 SubmitAction received");
      console.log("🟡 Action data:", JSON.stringify(action.data, null, 2));

      const { title, description, assignee, location, startDate, dueDate, status } = action.data;
      if (!title || !description || !assignee || !location || !status) {
        throw new Error("Missing one or more required fields");
      }
      const statusMap = {
        "Active": "Active",
        "Planned": "Deferred",
        "Completed": "Completed"
      };

      const wrikeStatus = statusMap[status];
      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;

      const wrikeResponse = await axios.post('https://www.wrike.com/api/v4/tasks', {
        title,
        description,
        importance: 'High',
        status: wrikeStatus,
        dates: {
          start: startDate,
          due: dueDate,
        },
        responsibles: [assignee],
        parents: [location],
      }, {
        headers: {
          Authorization: `Bearer ${wrikeToken}`,
          'Content-Type': 'application/json',
        },
      });

      const taskLink = wrikeResponse.data.data[0].permalink;
      const users = await this.fetchWrikeUsers();
      const assigneeDetails = users.find(u => u.id === assignee);
      const assigneeName = assigneeDetails ? assigneeDetails.name : assignee;

      const formattedDueDate = new Date(dueDate).toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric'
      });

      return {
        task: {
          type: 'continue',
          value: {
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard', version: '1.5',
              body: [
                { type: 'TextBlock', text: '🎉 Task Created Successfully!', weight: 'Bolder', size: 'Large', color: 'Good', wrap: true },
                { type: 'TextBlock', text: `**${title}**`, size: 'Medium', wrap: true },
                {
                  type: 'ColumnSet',
                  columns: [
                    {
                      type: 'Column', width: 'stretch', items: [
                        { type: 'TextBlock', text: `👤 **Assignee:** ${assigneeName}`, wrap: true },
                        { type: 'TextBlock', text: `📅 **Due Date:** ${formattedDueDate}`, wrap: true, spacing: 'Small' }
                      ]
                    }
                  ]
                }
              ],
              actions: [
                { type: 'Action.OpenUrl', title: '🔗 View Task in Wrike', url: taskLink }
              ]
            }),
            title: 'Task Created', height: 300, width: 450
          }
        }
      };
    } catch (error) {
      console.error("❌ Error in submitAction:", error.response?.data || error.message);
      return {
        task: { type: "message", value: `⚠️ Failed to create task: ${error.message}` }
      };
    }
  }

  async fetchWrikeUsers() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;

    try {
      const wrikeResponse = await axios.get('https://www.wrike.com/api/v4/contacts', {
        params: { deleted: false },
        headers: { Authorization: `Bearer ${wrikeToken}` },
      });

      const wrikeUsers = wrikeResponse.data.data;
      const graphUsers = await this.fetchGraphUsers();

      console.log("🟢 Wrike users:", wrikeUsers.length);
      wrikeUsers.forEach(u => {
        const email = u.profiles?.[0]?.email;
        console.log(`📨 Wrike user: ${email}`);
      });

      console.log("🟢 Graph users:", graphUsers.length);
      graphUsers.forEach(g => {
        console.log(`👥 Graph user: ${g.mail || g.userPrincipalName}`);
      });

      const acceptedDomains = ['@kashida.com', '@kashida-learning.com'];

      const validEmails = new Set(
        graphUsers
          .map(u => (u.mail || u.userPrincipalName)?.toLowerCase())
          .filter(email => acceptedDomains.some(domain => email?.endsWith(domain)))
      );

      const matched = wrikeUsers.filter(w => {
        const wrikeEmail = w.profiles?.[0]?.email?.toLowerCase();
        return wrikeEmail && validEmails.has(wrikeEmail);
      }).map(w => ({
        id: w.id,
        name: `${w.firstName || ''} ${w.lastName || ''}`.trim() + ` (${w.profiles[0]?.email})`
      }));

      console.log("✅ Matched Wrike + Graph users:", matched.length);
      return matched.length ? matched : [{ id: 'fallback', name: 'Fallback User' }];

    } catch (err) {
      console.error("❌ Matching error:", err?.response?.data || err.message);
      return [{ id: 'fallback', name: 'Fallback User' }];
    }
  }

  async fetchGraphUsers() {
    try {
      const tokenResponse = await cca.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default']
      });

      const accessToken = tokenResponse.accessToken;

      const response = await axios.get('https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName', {
        headers: { Authorization: `Bearer ${accessToken}` }
      });

      console.log("✅ Graph API returned users:", response.data.value.length);
      return response.data.value;
    } catch (err) {
      console.error("❌ Error in fetchGraphUsers:", err?.response?.data || err.message);
      return [];
    }
  }

  async fetchWrikeProjects() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: { Authorization: `Bearer ${wrikeToken}` },
    });
    return response.data.data.map(f => ({ id: f.id, title: f.title }));
  }
}

const bot = new WrikeBot();

server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

server.get('/', (req, res, next) => {
  res.send(200, '✔️ Railway bot is running!');
  return next();
});

server.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.send(400, 'Missing code from Wrike');

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

    console.log("🟢 Wrike Access Token:", response.data.access_token);
    res.send(200, "✅ Wrike authorization successful. You can close this.");
  } catch (error) {
    console.error("❌ Wrike OAuth Error:", error?.response?.data || error.message);
    res.send(500, "⚠️ Failed to authorize with Wrike.");
  }
});
