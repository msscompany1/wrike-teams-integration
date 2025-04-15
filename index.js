require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory } = require('botbuilder');

const PORT = process.env.PORT || 3978;
const server = restify.createServer();
server.listen(PORT, () => console.log(`✅ Bot is listening on http://localhost:${PORT}`));

const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context) {
    const messageHtml = context.activity.messagePayload?.body?.content || '';
    const plainText = messageHtml.replace(/<[^>]+>/g, '').trim();

    const cardPath = path.join(__dirname, 'cards', 'taskFormCard.json');
    const cardJson = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

    const descField = cardJson.body.find(f => f.id === 'description');
    if (descField) descField.value = plainText;

    const users = await this.fetchWrikeUsers();
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) userDropdown.choices = users.map(u => ({
      title: u.name,
      value: u.id
    }));

    const folders = await this.fetchWrikeProjects();
    const locationDropdown = cardJson.body.find(f => f.id === 'location');
    if (locationDropdown) locationDropdown.choices = folders.map(f => ({
      title: f.title,
      value: f.id
    }));

    const statuses = await this.fetchWrikeCustomStatuses();
    const statusDropdown = cardJson.body.find(f => f.id === 'status');
    if (statusDropdown) statusDropdown.choices = statuses.map(s => ({
      title: s.name,
      value: s.id
    }));

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
      console.log("🟡 Action data:", JSON.stringify(action.data, null, 2));
      const { title, description, assignee, location, startDate, dueDate, status } = action.data;

      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
      const taskPayload = {
        title,
        description,
        importance: 'High',
        customStatusId: status,
        ...(startDate || dueDate ? { dates: {} } : {}),
        ...(startDate ? { dates: { ...taskPayload.dates, start: startDate } } : {}),
        ...(dueDate ? { dates: { ...taskPayload.dates, due: dueDate } } : {}),
        ...(assignee ? { responsibles: [assignee] } : {}),
        ...(location ? { parents: [location] } : {}),
      };

      const response = await axios.post('https://www.wrike.com/api/v4/tasks', taskPayload, {
        headers: {
          Authorization: `Bearer ${wrikeToken}`,
          'Content-Type': 'application/json'
        }
      });

      const task = response.data.data[0];
      const link = task.permalink;

      return {
        task: {
          type: 'continue',
          value: {
            card: CardFactory.adaptiveCard({
              type: "AdaptiveCard",
              body: [
                {
                  type: "TextBlock",
                  size: "Large",
                  weight: "Bolder",
                  text: "✅ Wrike Task Created",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: `**${task.title}** has been created successfully.`,
                  wrap: true
                }
              ],
              actions: [
                {
                  type: "Action.OpenUrl",
                  title: "🔗 View in Wrike",
                  url: link
                }
              ],
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              version: "1.5"
            }),
          }
        }
      };
    } catch (err) {
      console.error("❌ Error in submitAction:", err.response?.data || err.message);
      return {
        task: {
          type: 'message',
          value: `⚠️ Failed to create task: ${err.message}`
        }
      };
    }
  }

  async fetchWrikeUsers() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const res = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: { Authorization: `Bearer ${wrikeToken}` }
    });
    return res.data.data.filter(u => !u.deleted).map(u => ({
      id: u.id,
      name: `${u.firstName} ${u.lastName}`
    }));
  }

  async fetchWrikeProjects() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const res = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: { Authorization: `Bearer ${wrikeToken}` }
    });
    return res.data.data.map(f => ({ id: f.id, title: f.title }));
  }

  async fetchWrikeCustomStatuses() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const res = await axios.get('https://www.wrike.com/api/v4/custom-statuses', {
      headers: { Authorization: `Bearer ${wrikeToken}` }
    });
    return res.data.data.map(s => ({ id: s.id, name: s.name }));
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
    const result = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        grant_type: 'authorization_code',
        code,
        redirect_uri: process.env.WRIKE_REDIRECT_URI,
      },
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });

    console.log("🟢 Wrike Access Token:", result.data.access_token);
    res.send(200, '✅ Authorization successful. You may close this window.');
  } catch (error) {
    console.error("❌ Wrike OAuth Error:", error.response?.data || error.message);
    res.send(500, '❌ Authorization failed.');
  }
});
