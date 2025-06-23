// Load environment variables
require('dotenv').config();

// Validate required environment variables
[
  'MICROSOFT_APP_ID',
  'MICROSOFT_APP_PASSWORD',
  'WRIKE_CLIENT_ID',
  'WRIKE_CLIENT_SECRET',
  'WRIKE_REDIRECT_URI',
  'TENANT_ID'
].forEach(key => {
  if (!process.env[key]) {
    console.error(`❌ Missing required environment variable: ${key}`);
    process.exit(1);
  }
});

// Global crash protection for production
process.on('unhandledRejection', reason => {
  console.error('💥 Unhandled Promise Rejection:', reason);
});
process.on('uncaughtException', err => {
  console.error('💥 Uncaught Exception:', err);
  process.exit(1);
});

const restify = require('restify');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const net = require('net');
const {
  BotFrameworkAdapter,
  MemoryStorage,
  ConversationState,
  TeamsActivityHandler,
  CardFactory
} = require('botbuilder');
const msal = require('@azure/msal-node');

// Load kill-port only after validating env
let killPort;
try {
  killPort = require('kill-port');
} catch {
  console.warn('⚠️ kill-port not installed. Skipping port cleanup.');
}

const PORT = process.env.PORT || 3978;

// HTTPS certificate loading with graceful error handling
let httpsOptions;
try {
  httpsOptions = {
    key: fs.readFileSync('/home/ubuntu/ssl/privkey.pem'),
    cert: fs.readFileSync('/home/ubuntu/ssl/fullchain.pem')
  };
} catch (err) {
  console.error('❌ SSL cert load failed:', err);
  process.exit(1);
}

const server = restify.createServer(httpsOptions);
server.use(restify.plugins.bodyParser());

// Bot adapter setup
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Catch-all for bot errors
adapter.onTurnError = async (context, error) => {
  console.error('❌ onTurnError:', error);
  await context.sendActivity('⚠️ Oops, something went wrong. Please try again later.');
};

// State storage
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

// MSAL config
const msalConfig = {
  auth: {
    clientId: process.env.MICROSOFT_APP_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.MICROSOFT_APP_PASSWORD
  }
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

// Teams bot implementation
class WrikeBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionFetchTask(context) {
    const messageHtml = context.activity.messagePayload?.body?.content || '';
    const plainTextMessage = messageHtml.replace(/<[^>]+>/g, '').trim();

    // Load Adaptive Card JSON template
    const cardJson = require('./taskFormCard.json');

    // Auto-fill title and description
    const titleField = cardJson.body.find(f => f.id === 'title');
    if (titleField) {
      titleField.value = plainTextMessage.slice(0, 50);
    }
    const descriptionField = cardJson.body.find(f => f.id === 'description');
    if (descriptionField) {
      descriptionField.value = plainTextMessage;
    }

    // Populate assignee dropdown
    const users = await this.fetchWrikeUsers();
    const userDropdown = cardJson.body.find(f => f.id === 'assignee');
    if (userDropdown) {
      userDropdown.choices = users.map(user => ({ title: user.name, value: user.id }));
    }

    // Populate location dropdown
    const folders = await this.fetchWrikeProjects();
    const locationDropdown = cardJson.body.find(f => f.id === 'location');
    if (locationDropdown) {
      locationDropdown.choices = folders.map(folder => ({ title: folder.title, value: folder.id }));
    }

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
    try {
      console.log('🔁 SubmitAction received');
      console.log('🟡 Action data:', JSON.stringify(action.data, null, 2));

      const { title, description, assignee, location, startDate, dueDate, status } = action.data;
      if (!title || !description || !assignee || !location || !status) {
        throw new Error('Missing one or more required fields');
      }

      const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
      const wrikeResponse = await axios.post('https://www.wrike.com/api/v4/tasks', {
        title,
        description,
        importance: 'High',
        startDate,
        dueDate,
        status,
        responsibles: [assignee],
        parents: [location]
      }, {
        headers: {
          Authorization: `Bearer ${wrikeToken}`,
          'Content-Type': 'application/json'
        }
      });

      const taskLink = wrikeResponse.data.data[0].permalink;
      console.log('✅ Wrike Task Created:', taskLink);

      return {
        task: {
          type: 'message',
          value: `✅ Task created in Wrike: [${title}](${taskLink})`
        }
      };
    } catch (error) {
      console.error('❌ Error in submitAction:', error.response?.data || error.message);
      return {
        task: {
          type: 'message',
          value: `⚠️ Failed to create task: ${error.message}`
        }
      };
    }
  }

  async fetchWrikeUsers() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/contacts', {
      headers: {
        Authorization: `Bearer ${wrikeToken}`
      }
    });
    return response.data.data
      .filter(u => !u.deleted)
      .map(u => ({ id: u.id, name: `${u.firstName} ${u.lastName}`.trim() }));
  }

  async fetchWrikeProjects() {
    const wrikeToken = process.env.WRIKE_ACCESS_TOKEN;
    const response = await axios.get('https://www.wrike.com/api/v4/folders?project=true', {
      headers: {
        Authorization: `Bearer ${wrikeToken}`
      }
    });
    return response.data.data.map(f => ({ id: f.id, title: f.title }));
  }
}

const bot = new WrikeBot();

// Main message endpoint
server.post('/api/messages', async (req, res) => {
  await adapter.processActivity(req, res, async context => {
    await bot.run(context, conversationState);
  });
});

// Health check route
server.get('/', (req, res, next) => {
  res.send(200, '✔️ Railway bot is running!');
  return next();
});

// OAuth callback route
server.get('/auth/callback', async (req, res) => {
  const { code, state: userId } = req.query;
  if (!code) {
    res.send(400, 'Missing code from Wrike');
    return;
  }
  try {
    const tokenResponse = await axios.post('https://login.wrike.com/oauth2/token', {
      code,
      grant_type: 'authorization_code',
      client_id: process.env.WRIKE_CLIENT_ID,
      client_secret: process.env.WRIKE_CLIENT_SECRET,
      redirect_uri: process.env.WRIKE_REDIRECT_URI
    });
    // Persist token (consider moving to a database for production)
    // wrikeTokens.set(userId, tokenResponse.data.access_token);
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`<html><body style="text-align:center;font-family:Arial,sans-serif;">Authenticated!<br><a href="${process.env.TEAMS_LINK_CUSTOM_FIELD_ID||'#'}">Return to Teams</a></body></html>`);
  } catch (err) {
    console.error('❌ OAuth Callback Error:', err.response?.data || err.message);
    res.send(500, '❌ Authorization failed');
  }
});

// Function to check port availability
const checkPort = port => new Promise((resolve, reject) => {
  const tester = net.createServer()
    .once('error', err => (err.code === 'EADDRINUSE' ? reject(err) : resolve()))
    .once('listening', () => tester.once('close', () => resolve()).close())
    .listen(port);
});

// Sequential startup: clear port, check, then start server
let startSequence = Promise.resolve();
if (killPort) {
  startSequence = startSequence
    .then(() => killPort(PORT, 'tcp'))
    .then(() => console.log(`✅ Cleared port ${PORT} before startup`))
    .catch(err => console.warn(`⚠️ Could not clear port ${PORT}:`, err.message));
}

startSequence
  .then(() => checkPort(PORT))
  .then(() => {
    server.listen(PORT, () => console.log(`✅ HTTPS bot running on https://wrike-bot.kashida-learning.co:${PORT}`));
  })
  .catch(err => {
    console.error('❌ Startup failed:', err);
    process.exit(1);
  });
