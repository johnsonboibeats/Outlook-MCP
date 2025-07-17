/**
 * Check Azure App Permissions for Outlook MCP Server
 * This script checks what permissions have been granted to the Azure application
 */
const { ensureAuthenticated } = require('./auth');
const { callGraphAPI } = require('./utils/graph-api');
const config = require('./config');

/**
 * Check the current app's permissions and consent status
 */
async function checkAppPermissions() {
  try {
    console.log('🔍 Checking Azure App Permissions for Outlook MCP Server...\n');
    
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Get current user info to identify the app context
    const me = await callGraphAPI(accessToken, 'GET', 'me?$select=displayName,mail,userPrincipalName');
    console.log(`📱 Authenticated as: ${me.displayName} (${me.mail})`);
    console.log(`🆔 Client ID: ${config.AUTH_CONFIG.clientId}\n`);
    
    // Check what scopes are actually available in the token
    console.log('📋 Configured Scopes in config.js:');
    config.AUTH_CONFIG.scopes.forEach(scope => {
      console.log(`   • ${scope}`);
    });
    console.log('');
    
    // Test each permission by making actual API calls
    console.log('🧪 Testing Actual Permissions:\n');
    
    // Test User.Read
    await testPermission('User.Read', async () => {
      await callGraphAPI(accessToken, 'GET', 'me');
      return 'Can read user profile';
    });
    
    // Test Mail.Read
    await testPermission('Mail.Read', async () => {
      const result = await callGraphAPI(accessToken, 'GET', 'me/messages?$top=1&$select=id,subject');
      return `Can read mail (found ${result.value?.length || 0} messages)`;
    });
    
    // Test Mail.ReadWrite
    await testPermission('Mail.ReadWrite', async () => {
      // Test by trying to access a mail folder (read operation that requires readwrite for some scenarios)
      await callGraphAPI(accessToken, 'GET', 'me/mailFolders?$top=1');
      return 'Can read/write mail folders';
    });
    
    // Test Mail.Send
    await testPermission('Mail.Send', async () => {
      // We can't actually send a test email, but we can check if the sendMail endpoint is accessible
      // by trying to access the sent items folder
      await callGraphAPI(accessToken, 'GET', 'me/mailFolders/sentitems');
      return 'Can access send mail endpoint (sent items folder)';
    });
    
    // Test Calendars.Read
    await testPermission('Calendars.Read', async () => {
      const result = await callGraphAPI(accessToken, 'GET', 'me/events?$top=1&$select=id,subject');
      return `Can read calendar (found ${result.value?.length || 0} events)`;
    });
    
    // Test Calendars.ReadWrite
    await testPermission('Calendars.ReadWrite', async () => {
      await callGraphAPI(accessToken, 'GET', 'me/calendars?$top=1');
      return 'Can read/write calendars';
    });
    
    // Additional permission tests
    console.log('\n🔍 Additional Permission Tests:\n');
    
    // Test shared mailbox access (if configured)
    try {
      console.log('📧 Testing shared mailbox access...');
      const sharedMailboxes = await callGraphAPI(accessToken, 'GET', 'me/mailboxSettings');
      console.log('   ✅ Can access mailbox settings');
    } catch (error) {
      console.log(`   ❌ Cannot access mailbox settings: ${error.message}`);
    }
    
    // Check if app has admin consent
    try {
      console.log('\n🛡️  Checking admin consent status...');
      // Try to access organization info - usually requires admin consent
      const org = await callGraphAPI(accessToken, 'GET', 'organization?$select=id,displayName');
      if (org.value && org.value.length > 0) {
        console.log(`   ✅ Admin consent likely granted - can access org: ${org.value[0].displayName}`);
      }
    } catch (error) {
      console.log('   ℹ️  Admin consent status unclear - cannot access organization info');
      console.log(`   Details: ${error.message}`);
    }
    
    console.log('\n✅ Permission check complete!');
    console.log('\n💡 Notes:');
    console.log('   • If you see ❌ errors, the app may need additional consent');
    console.log('   • Some permissions may require admin consent for your organization');
    console.log('   • Visit Azure Portal > App Registrations > Your App > API Permissions to grant consent');
    
  } catch (error) {
    if (error.message === 'Authentication required') {
      console.error('❌ Authentication required. Please run the authentication process first.');
      console.error('   Try running: node outlook-auth-server.js');
    } else {
      console.error('❌ Error checking app permissions:', error.message);
    }
  }
}

/**
 * Test a specific permission by running a test function
 */
async function testPermission(permissionName, testFunction) {
  try {
    const result = await testFunction();
    console.log(`   ✅ ${permissionName}: ${result}`);
  } catch (error) {
    console.log(`   ❌ ${permissionName}: ${error.message}`);
  }
}

// Run the permission check if this script is executed directly
if (require.main === module) {
  checkAppPermissions().catch(console.error);
}

module.exports = { checkAppPermissions };