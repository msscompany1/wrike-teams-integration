require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const {
  BotFrameworkAdapter,
  MemoryStorage,
  ConversationState,
} = require('botbuilder');
const {
  TeamsActivityHandler,
  CardFactory,
} = require('botbuilder');
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
    const messageHtml = context.activity.messagePayload?.body?.content || '';
    const plainTextMessage = messageHtml.replace(/<[^>]+>/g, '').trim();
    console.log("🟡 Raw message HTML:", messageHtml);
    console.log("🟡 Plain text extracted:", plainTextMessage);
    
    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));
  
    const descriptionField = cardJson.body.find(f => f.id === 'description');
    if (descriptionField) {
      descriptionField.value = plainTextMessage;
    }
    console.log("🟡 context.activity.messagePayload:", context.activity.messagePayload);

    const users = await this.fetchWrikeUsers();
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) {
      userDropdown.choices = users.map(user => ({
        title: user.name,
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
          "Active": "Active",         // valid
          "Planned": "Deferred",      // mapped to valid Wrike status
          "Completed": "Completed"    // valid
      };

     const wrikeStatus = statusMap[status];
        if (!wrikeStatus) {
          throw new Error(`Unsupported status: ${status}`);
     }

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
      console.log("✅ Wrike Task Created:", taskLink);

      const users = await this.fetchWrikeUsers();
      const assigneeDetails = users.find(u => u.id === assignee);
      const assigneeName = assigneeDetails ? assigneeDetails.name : assignee;
      const assigneeAvatar = assigneeDetails?.photo?.url || 'https://static.wrike.com/favicon.ico';

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
                text: '🎉 Task Created Successfully!',
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
                type: 'ColumnSet',
                columns: [
                  {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                      {
                        type: 'TextBlock',
                        text: `👤 **Assignee:** ${assigneeName}`,
                        wrap: true
                      },
                      {
                        type: 'TextBlock',
                        text: `📅 **Due Date:** ${formattedDueDate}`,
                        wrap: true,
                        spacing: 'Small'
                      }
                    ]
                  }
                ]
              }
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
        task: {
          type: "message",
          value: `⚠️ Failed to create task: ${error.message}`
        }
      };
    }
  }

  async fetchWrikeUsers() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: {
        Authorization: `Bearer ${wrikeToken}`,
      },
    });
    return response.data.data
      .filter(u => !u.deleted)
      .map(u => ({ id: u.id, name: `${u.firstName} ${u.lastName}`.trim() }));
  }

  async fetchWrikeProjects() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: {
        Authorization: `Bearer ${wrikeToken}`,
      },
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
  if (!code) {
    res.send(400, 'Missing code from Wrike');
    return;
  }

  try {
    const response = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        grant_type: 'authorization_code',
        code,
        redirect_uri: process.env.WRIKE_REDIRECT_URI,
      },
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
    });

    console.log("🟢 Wrike Access Token:", response.data.access_token);
    res.send(200, "✅ Wrike authorization successful. You can close this.");
  } catch (error) {
    console.error("❌ Wrike OAuth Error:", error?.response?.data || error.message);
    res.send(500, "⚠️ Failed to authorize with Wrike.");
  }
});
