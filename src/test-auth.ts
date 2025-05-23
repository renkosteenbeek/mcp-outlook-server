import { AuthManager } from './auth.js';
import { GraphClient } from './graph-client.js';
import express from 'express';
import open from 'open';
import { config } from './config.js';

const app = express();
const authManager = new AuthManager();
let authCodeReceived = false;

// Setup callback route
app.get('/auth/callback', async (req, res) => {
  const code = req.query.code as string;
  
  if (!code) {
    res.send('Error: No authorization code received');
    return;
  }

  try {
    console.log('📥 Received authorization code');
    const tokenResponse = await authManager.acquireTokenByCode(code);
    console.log('✅ Token acquired successfully');
    console.log(`👤 Authenticated as: ${tokenResponse.account.username}`);
    
    authCodeReceived = true;
    
    res.send(`
      <html>
        <body>
          <h2>Authentication Successful!</h2>
          <p>You can close this window and return to the terminal.</p>
          <script>window.close();</script>
        </body>
      </html>
    `);
    
    // Run tests after successful auth
    setTimeout(async () => {
      await runTests();
      process.exit(0);
    }, 1000);
    
  } catch (error) {
    console.error('❌ Authentication error:', error);
    res.send('Authentication failed. Check the console for details.');
  }
});

async function runTests() {
  console.log('\n🧪 Starting Microsoft Graph API tests...\n');
  
  const graphClient = new GraphClient(authManager);
  
  try {
    // Test 1: Get user profile
    console.log('📋 Test 1: Getting user profile...');
    const user = await graphClient.getUser();
    console.log(`✅ User: ${user.displayName} (${user.mail || user.userPrincipalName})`);
    
    // Test 2: List calendars
    console.log('\n📋 Test 2: Listing calendars...');
    const calendars = await graphClient.listCalendars();
    console.log(`✅ Found ${calendars.value.length} calendar(s):`);
    calendars.value.forEach((cal: any) => {
      console.log(`   - ${cal.name} (${cal.isDefaultCalendar ? 'Default' : 'Additional'})`);
    });
    
    // Test 3: List recent calendar events
    console.log('\n📋 Test 3: Listing recent calendar events...');
    const events = await graphClient.listCalendarEvents();
    console.log(`✅ Found ${events.value.length} event(s) in the next 7 days:`);
    events.value.slice(0, 5).forEach((event: any) => {
      console.log(`   - ${event.subject} (${new Date(event.start.dateTime).toLocaleString()})`);
    });
    
    // Test 4: List mail folders
    console.log('\n📋 Test 4: Listing mail folders...');
    const folders = await graphClient.listMailFolders();
    console.log(`✅ Found ${folders.value.length} mail folder(s):`);
    folders.value.forEach((folder: any) => {
      console.log(`   - ${folder.displayName} (${folder.totalItemCount} items)`);
    });
    
    // Test 5: List recent emails
    console.log('\n📋 Test 5: Listing recent emails...');
    const messages = await graphClient.listMessages();
    console.log(`✅ Found ${messages.value.length} recent email(s):`);
    messages.value.slice(0, 5).forEach((msg: any) => {
      console.log(`   - ${msg.subject} (from: ${msg.from?.emailAddress?.address || 'Unknown'})`);
    });
    
    console.log('\n✅ All tests completed successfully!');
    
  } catch (error) {
    console.error('\n❌ Test failed:', error);
  }
}

async function startAuthFlow() {
  const server = app.listen(config.port, async () => {
    console.log(`🚀 Auth server listening on port ${config.port}`);
    
    try {
      const authUrl = await authManager.getAuthUrl();
      console.log('\n🔐 Opening browser for authentication...');
      console.log('📋 If browser doesn\'t open, visit this URL:');
      console.log(authUrl);
      
      await open(authUrl);
      
      // Wait for auth callback
      console.log('\n⏳ Waiting for authentication callback...');
      
    } catch (error) {
      console.error('❌ Failed to start auth flow:', error);
      process.exit(1);
    }
  });
  
  // Timeout after 5 minutes
  setTimeout(() => {
    if (!authCodeReceived) {
      console.error('\n❌ Authentication timeout. Please try again.');
      process.exit(1);
    }
  }, 300000);
}

// Check if we already have a valid token
async function checkExistingAuth() {
  const token = await authManager.getAccessToken();
  if (token) {
    console.log('✅ Found valid cached token');
    await runTests();
    process.exit(0);
  } else {
    console.log('🔐 No valid token found, starting authentication flow...');
    await startAuthFlow();
  }
}

// Start the test
checkExistingAuth();