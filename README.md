# MCP Outlook Server

Een MCP (Model Context Protocol) server voor Microsoft Outlook integratie met multi-tenant support.

## Functies

- 🔐 OAuth2 authenticatie met Microsoft
- 📅 Agenda beheer (lezen, schrijven, events maken)
- 📧 Email beheer (lezen, verzenden, folders)
- 🏢 Multi-tenant support (meerdere Azure tenants)
- 💾 Automatische token refresh

## Installatie

```bash
# Clone de repository
git clone <repository-url>
cd mcp-outlook-server

# Installeer dependencies
npm install

# Build het project
npm run build
```

## Configuratie

### 1. Azure App Registration

Voor elke tenant heb je een App Registration nodig in Azure Portal:

1. Ga naar [Azure Portal](https://portal.azure.com)
2. Ga naar "App registrations" → "New registration"
3. Configureer:
   - Name: `MCP Outlook Integration`
   - Supported account types: "Accounts in any organizational directory"
   - Redirect URI: Web → `http://localhost:3000/auth/callback`
4. Ga naar "Certificates & secrets" → "New client secret"
5. Ga naar "API permissions" → "Add a permission" → "Microsoft Graph":
   - Delegated permissions:
     - `User.Read`
     - `Mail.Read`
     - `Mail.ReadWrite`
     - `Mail.Send`
     - `Calendars.Read`
     - `Calendars.ReadWrite`

### 2. Environment Variables

Maak een `.env` file:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
TEST_USER_EMAIL=your-email@domain.com
REDIRECT_URI=http://localhost:3000/auth/callback
PORT=3000
```

### 3. Claude Desktop Configuratie

Voeg toe aan je Claude Desktop config (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/path/to/mcp-outlook-server/dist/index.js"],
      "env": {
        "TENANT_ID": "your-tenant-id",
        "CLIENT_ID": "your-client-id",
        "CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## Gebruik in Claude

### Authenticatie
```
Gebruik de outlook_auth_login tool om in te loggen
```

### Agenda functies
```
- Lijst agenda's: outlook_calendar_list
- Bekijk events: outlook_calendar_events_list
- Maak event: outlook_calendar_event_create met subject, startDateTime, endDateTime
```

### Email functies
```
- Lijst folders: outlook_mail_folders_list
- Bekijk emails: outlook_mail_messages_list
- Verstuur email: outlook_mail_send met to, subject, body
```

## Development

```bash
# Run in development mode
npm run dev

# Test authentication
npm run test
```

## Multi-Tenant Support

Voor multi-tenant support is uitbreiding nodig. Huidige versie werkt met één tenant tegelijk.

## Security

- Client secrets worden veilig opgeslagen in environment variables
- Tokens worden encrypted opgeslagen
- Automatische token refresh
- OAuth2 met PKCE voor extra beveiliging

## Troubleshooting

1. **"No reply address registered"**: Voeg redirect URI toe in Azure Portal
2. **"Invalid client secret"**: Controleer of secret niet verlopen is
3. **"Permission denied"**: Controleer API permissions in Azure Portal
4. **Token expired**: Server refresht automatisch, probeer opnieuw