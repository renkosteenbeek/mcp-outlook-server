[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/renkosteenbeek-mcp-outlook-server-badge.png)](https://mseep.ai/app/renkosteenbeek-mcp-outlook-server)

# MCP Outlook Server

An MCP (Model Context Protocol) server for Microsoft Outlook integration with multi-account support.

## Features

- 🔐 OAuth2 authentication with Microsoft
- 📅 Calendar management (read, write, create events)
- 📧 Email management (read, send, folders)
- 🏢 Multi-account support (multiple Microsoft accounts)
- 💾 Automatic token refresh
- 🔄 Concurrent operations across all accounts

## Installation

```bash
# Clone the repository
git clone https://github.com/renkosteenbeek/mcp-outlook-server.git
cd mcp-outlook-server

# Install dependencies
npm install

# Build the project
npm run build
```

## Configuration

### 1. Azure App Registration

For each Microsoft account you want to connect, you need an App Registration.

#### Direct link to App Registrations:
[https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)

#### Steps to create an App Registration:

1. Go to the [App Registrations page](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click "New registration"
3. Configure the application:
   - **Name**: `MCP Outlook Integration` (or any name you prefer)
   - **Supported account types**: Choose based on your needs:
     - For personal Microsoft accounts: "Accounts in any organizational directory and personal Microsoft accounts"
     - For work/school accounts only: "Accounts in this organizational directory only"
     - For multi-tenant: "Accounts in any organizational directory"
   - **Redirect URI**: 
     - Platform: `Web`
     - URI: `http://localhost:3000/auth/callback`
   - Click "Register"

4. After registration, save these values:
   - **Application (client) ID**: Found on the Overview page
   - **Directory (tenant) ID**: Found on the Overview page (use "common" for personal accounts)

5. Create a client secret:
   - Go to "Certificates & secrets" in the left menu
   - Click "New client secret"
   - Add a description and select expiration period
   - Click "Add"
   - **IMPORTANT**: Copy the secret value immediately (it won't be shown again)

6. Configure API permissions:
   - Go to "API permissions" in the left menu
   - Click "Add a permission"
   - Select "Microsoft Graph"
   - Choose "Delegated permissions"
   - Select these permissions:
     - `User.Read`
     - `Mail.Read`
     - `Mail.ReadWrite`
     - `Mail.Send`
     - `Calendars.Read`
     - `Calendars.ReadWrite`
   - Click "Add permissions"
   - If you're an admin, click "Grant admin consent" (optional)

7. Configure Authentication settings:
   - Go to "Authentication" in the left menu
   - Under "Platform configurations", ensure you have a "Web" platform
   - Verify the redirect URI is: `http://localhost:3000/auth/callback`
   - Under "Implicit grant and hybrid flows", leave both options unchecked
   - Click "Save"

### 2. Configuration File

Create a `config.json` file in the project root:

```json
{
  "accounts": [
    {
      "name": "Personal",
      "tenantId": "common",
      "clientId": "your-client-id",
      "clientSecret": "your-client-secret"
    },
    {
      "name": "Work",
      "tenantId": "your-tenant-id",
      "clientId": "your-client-id",
      "clientSecret": "your-client-secret"
    }
  ],
  "server": {
    "port": 3000,
    "redirectUri": "http://localhost:3000/auth/callback"
  }
}
```

### 3. Claude Desktop Configuration

Add to your Claude Desktop config (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/path/to/mcp-outlook-server/dist/index.js"],
      "env": {}
    }
  }
}
```

## Usage in Claude

### Authentication
```
Login to your accounts: outlook_auth_login
```

### Calendar Functions
```
- List calendars: outlook_calendar_list
- View events: outlook_calendar_events_list
- Create event: outlook_calendar_event_create with subject, startDateTime, endDateTime
```

### Email Functions
```
- List folders: outlook_mail_folders_list
- View emails: outlook_mail_messages_list
- Send email: outlook_mail_send with to, subject, body
```

### Multi-Account Support
- All operations run on all configured accounts by default
- Use the `account` parameter to target a specific account
- Responses include the account name for clarity

## Development

```bash
# Run in development mode
npm run dev

# Test authentication
npm run test
```

## Multi-Account Features

- Configure multiple Microsoft accounts (personal, work, etc.)
- Operations execute across all accounts simultaneously
- Filter results by account name
- Each account maintains its own authentication state

## Security

- Client secrets are stored securely in config files
- Tokens are stored encrypted per account
- Automatic token refresh per account
- OAuth2 with PKCE for enhanced security

## Troubleshooting

1. **"No reply address registered"**: Add redirect URI in Azure Portal
2. **"Invalid client secret"**: Check if secret has expired
3. **"Permission denied"**: Verify API permissions in Azure Portal
4. **Token expired**: Server refreshes automatically, try again