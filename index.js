// Enhanced index.js: icons in labels, smart defaults, importance emojis, and assignee emoji prefix
require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');
const msal = require('@azure/msal-node');

const PORT = process.env.PORT || 3978;
const CUSTOM_FIELD_ID_TEAMS_LINK = process.env.TEAMS_LINK_CUSTOM_FIELD_ID;

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

    // Autofill description
    const descriptionField = cardJson.body.find(f => f.id === 'description');
    if (descriptionField) {
      descriptionField.value = plainTextMessage;
    }

    // Autofill due date = today + 1
    const dueDateField = cardJson.body.find(f => f.id === 'dueDate');
    if (dueDateField) {
      const tomorrow = new Date();
      tomorrow.setDate(tomorrow.getDate() + 1);
      dueDateField.value = tomorrow.toISOString().split('T')[0];
    }

    // Pre-select Normal importance by default
    const importanceField = cardJson.body.find(f => f.id === 'importance');
    if (importanceField) {
      importanceField.value = 'Normal';
      importanceField.choices = [
        { title: '🔴 High', value: 'High' },
        { title: '🟡 Normal', value: 'Normal' },
        { title: '🟢 Low', value: 'Low' }
      ];
    }

    // Enhance assignee list with emoji
    const users = await this.fetchWrikeUsers();
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) {
      userDropdown.choices = users.map(user => ({
        title: `${user.name}`,
        value: user.id,
      }));
    }

    // Populate locations
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
      const { title, description, assignee, location, startDate, dueDate, importance, comment } = action.data;
      if (!title || !description || !assignee || !location || !importance) {
        throw new Error("Missing one or more required fields");
      }

      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
      const assigneeIds = Array.isArray(assignee) ? assignee : [assignee];
      const teamsMessageLink = context.activity.value?.messagePayload?.linkToMessage || '';

      const wrikeResponse = await axios.post('https://www.wrike.com/api/v4/tasks', {
        title,
        description,
        importance,
        status: "Active",
        dates: { start: startDate, due: dueDate },
        responsibles: assigneeIds,
        parents: [location],
        customFields: [
          { id: CUSTOM_FIELD_ID_TEAMS_LINK, value: teamsMessageLink }
        ]
      }, {
        headers: {
          Authorization: `Bearer ${wrikeToken}`,
          'Content-Type': 'application/json',
        },
      });

      const createdTaskId = wrikeResponse.data.data[0].id;
      const taskLink = wrikeResponse.data.data[0].permalink;
      const users = await this.fetchWrikeUsers();
      const assigneeNames = assigneeIds.map(id => {
        const user = users.find(u => u.id === id);
        return user ? user.name : id;
      });

      if (comment) {
        await axios.post(`https://www.wrike.com/api/v4/tasks/${createdTaskId}/comments`, {
          text: comment
        }, {
          headers: {
            Authorization: `Bearer ${wrikeToken}`,
            'Content-Type': 'application/json'
          }
        });
      }

      const formattedDueDate = new Date(dueDate).toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric'
      });
      
      return {
        task: {
          type: 'continue',
          value: {
            card: CardFactory.adaptiveCard({
              type: 'AdaptiveCard',
              version: '1.5',
              body: [
                {
                  type: 'TextBlock',
                  text: '✅ Task Created Successfully!',
                  weight: 'Bolder',
                  size: 'Large',
                  color: 'Good',
                  wrap: true
                },
                {
                  type: 'TextBlock',
                  text: `**${title}**`,
                  size: 'Medium',
                  wrap: true
                },
                {
                  type: 'TextBlock',
                  text: '📌 Task Details',
                  weight: 'Bolder',
                  color: 'Accent',
                  spacing: 'Medium'
                },
                {
                  type: 'ColumnSet',
                  columns: [
                    {
                      type: 'Column',
                      width: 'stretch',
                      items: [
                        {
                          type: 'TextBlock',
                          text: `👥 **Assignees:** ${assigneeNames.join(', ')}`,
                          wrap: true
                        },
                        {
                          type: 'TextBlock',
                          text: `📅 **Due Date:** ${formattedDueDate}`,
                          wrap: true,
                          spacing: 'Small'
                        },
                        {
                          type: 'TextBlock',
                          text: `📊 **Importance:** ${importance}`,
                          wrap: true,
                          spacing: 'Small'
                        }
                      ]
                    }
                  ]
                },
             
              ],
              actions: [
                {
                  type: 'Action.OpenUrl',
                  title: '🔗 View Task in Wrike',
                  url: taskLink
                }
              ]
            }),
            title: 'Task Created',
            height: 300,
            width: 450
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
      return wrikeUsers.filter(w => {
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

  async fetchWrikeProjects() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: { Authorization: `Bearer ${wrikeToken}` },
    });
    return response.data.data.map(f => ({ id: f.id, title: f.title }));
  }

  async fetchGraphUsers() {
    return [];
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
