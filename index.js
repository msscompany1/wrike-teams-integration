const fs = require('fs');
const path = require('path');
const https = require('https');
const net = require('net');
const restify = require('restify');
const axios = require('axios');
const { BotFrameworkAdapter, MemoryStorage, ConversationState } = require('botbuilder');
const { TeamsActivityHandler, CardFactory, TurnContext } = require('botbuilder');

require('dotenv').config();
const { saveTokens, loadTokens } = require('./wrike-db');
const wrikeDB = require('./wrike-db'); // ✅ Needed for token saving

const PORT = process.env.PORT || 3978;
const wrikeTokens = new Map();
const CUSTOM_FIELD_ID_TEAMS_LINK = process.env.TEAMS_LINK_CUSTOM_FIELD_ID;

// SSL for production
const httpsOptions = {
  key: fs.readFileSync('/home/ubuntu/ssl/privkey.pem'),
  cert: fs.readFileSync('/home/ubuntu/ssl/fullchain.pem')
};

const server = restify.createServer(httpsOptions);
server.use(restify.plugins.queryParser());

const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD
});

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

// Graceful error logging
process.on('unhandledRejection', (reason) => {
  console.error('💥 Unhandled Promise Rejection:', reason);
});
process.on('uncaughtException', (err) => {
  console.error('💥 Uncaught Exception:', err);
});

try {
  require('kill-port')(PORT, 'tcp')
    .then(() => console.log(`✅ Cleared port ${PORT} before startup`))
    .catch((err) => console.warn('⚠️ Could not clear port:', err.message));
} catch (e) {
  console.warn('⚠️ kill-port not installed. Skipping port cleanup.');
}

const checkPort = (port) => new Promise((resolve, reject) => {
  const tester = net.createServer()
    .once('error', err => (err.code === 'EADDRINUSE' ? reject(err) : resolve()))
    .once('listening', () => tester.once('close', () => resolve()).close())
    .listen(port);
});

checkPort(PORT)
  .then(() => server.listen(PORT, () => {
    console.log(`✅ HTTPS bot running on https://wrike-bot.kashida-learning.co:${PORT}`);
  }))
  .catch(err => {
    console.error(`❌ Port ${PORT} in use:`, err.message);
    process.exit(1);
  });

class WrikeBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      await context.sendActivity('👋 Please use the messaging extension to create a Wrike task.');
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let member of membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity('👋 Welcome to the Wrike Task Bot!');
        }
      }
      await next();
    });

    this.onInvokeActivity = async (context) => {
      const invoke = context.activity;
      const userId = invoke.from.aadObjectId;

      // Load from DB if not in memory
      if (!wrikeTokens.has(userId)) {
        await new Promise((resolve) => {
          wrikeDB.loadTokens(userId, (tokenData) => {
            if (tokenData) {
              wrikeTokens.set(userId, tokenData);
            }
            resolve();
          });
        });
      }

      if (invoke.name === 'composeExtension/query') {
        const tokenInfo = wrikeTokens.get(userId);
        const now = Date.now();
        if (!tokenInfo || tokenInfo.expiresAt < now) {
          return {
            composeExtension: {
              type: 'auth',
              suggestedActions: {
                actions: [
                  {
                    type: 'openUrl',
                    value: `https://wrike-bot.kashida-learning.co/auth/callback?state=${userId}`,
                    title: 'Login to Wrike'
                  }
                ]
              }
            }
          };
        }

        // Return the compose card UI
        return {
          composeExtension: {
            type: 'botMessagePreview',
            activityPreview: {
              type: 'message',
              attachments: [
                CardFactory.heroCard('Wrike Task Creation', 'Click "Create Task" to start.')
              ]
            }
          }
        };
      }

      return { status: 200 };
    };
  }
}

const bot = new WrikeBot();

server.post('/api/messages', (req, res, next) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  }).then(() => next()).catch(err => {
    console.error('💥 processActivity error:', err);
    next(err);
  });
});

server.get('/auth/callback', async (req, res) => {
  try {
    const { code, state: userId } = req.query;
    const tr = await axios.post('https://login.wrike.com/oauth2/token', null, {
      params: {
        grant_type: 'authorization_code',
        code,
        client_id: process.env.WRIKE_CLIENT_ID,
        client_secret: process.env.WRIKE_CLIENT_SECRET,
        redirect_uri: process.env.WRIKE_REDIRECT_URI
      }
    });

    const expiresAt = Date.now() + (tr.data.expires_in * 1000);

    // Save to both memory and SQLite DB
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
          <p style="margin-top:20px;">You may now return to Microsoft Teams to continue your task.</p>
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
