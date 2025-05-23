import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { AuthManager } from './auth.js';
import { GraphClient } from './graph-client.js';
import express from 'express';
import open from 'open';
import { config } from './config.js';

interface OutlookTool {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, any>;
    required?: string[];
  };
}

const tools: OutlookTool[] = [
  {
    name: 'outlook_auth_login',
    description: 'Authenticate with Microsoft Outlook. Opens a browser for OAuth2 authentication.',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'outlook_auth_status',
    description: 'Check authentication status and get current user info',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'outlook_calendar_list',
    description: 'List all calendars for the authenticated user',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'outlook_calendar_events_list',
    description: 'List calendar events within a date range',
    inputSchema: {
      type: 'object',
      properties: {
        calendarId: {
          type: 'string',
          description: 'Calendar ID (optional, defaults to primary calendar)',
        },
        startDateTime: {
          type: 'string',
          description: 'Start date/time in ISO format (optional, defaults to now)',
        },
        endDateTime: {
          type: 'string',
          description: 'End date/time in ISO format (optional, defaults to 7 days from now)',
        },
      },
    },
  },
  {
    name: 'outlook_calendar_event_create',
    description: 'Create a new calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        subject: {
          type: 'string',
          description: 'Event subject/title',
        },
        body: {
          type: 'string',
          description: 'Event body/description (optional)',
        },
        startDateTime: {
          type: 'string',
          description: 'Start date/time in ISO format',
        },
        endDateTime: {
          type: 'string',
          description: 'End date/time in ISO format',
        },
        location: {
          type: 'string',
          description: 'Event location (optional)',
        },
        attendees: {
          type: 'array',
          description: 'Array of attendee email addresses (optional)',
          items: {
            type: 'string',
          },
        },
        isOnlineMeeting: {
          type: 'boolean',
          description: 'Whether to create an online meeting (optional)',
        },
        calendarId: {
          type: 'string',
          description: 'Calendar ID (optional, defaults to primary calendar)',
        },
      },
      required: ['subject', 'startDateTime', 'endDateTime'],
    },
  },
  {
    name: 'outlook_mail_folders_list',
    description: 'List all mail folders',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'outlook_mail_messages_list',
    description: 'List mail messages',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: {
          type: 'string',
          description: 'Folder ID (optional, defaults to inbox)',
        },
        filter: {
          type: 'string',
          description: 'OData filter string (optional)',
        },
      },
    },
  },
  {
    name: 'outlook_mail_message_get',
    description: 'Get a specific email message by ID',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: {
          type: 'string',
          description: 'Message ID',
        },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'outlook_mail_send',
    description: 'Send an email message',
    inputSchema: {
      type: 'object',
      properties: {
        to: {
          type: 'array',
          description: 'Array of recipient email addresses',
          items: {
            type: 'string',
          },
        },
        subject: {
          type: 'string',
          description: 'Email subject',
        },
        body: {
          type: 'string',
          description: 'Email body (HTML supported)',
        },
        cc: {
          type: 'array',
          description: 'Array of CC recipient email addresses (optional)',
          items: {
            type: 'string',
          },
        },
        bcc: {
          type: 'array',
          description: 'Array of BCC recipient email addresses (optional)',
          items: {
            type: 'string',
          },
        },
        isHtml: {
          type: 'boolean',
          description: 'Whether the body is HTML (optional, defaults to false)',
        },
      },
      required: ['to', 'subject', 'body'],
    },
  },
];

class OutlookMCPServer {
  private server: Server;
  private authManager: AuthManager;
  private graphClient: GraphClient;
  private authServer: any = null;

  constructor() {
    this.server = new Server(
      {
        name: 'mcp-outlook-server',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.authManager = new AuthManager();
    this.graphClient = new GraphClient(this.authManager);

    this.setupHandlers();

    this.server.onerror = (error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.cleanup();
      process.exit(0);
    });
  }

  private setupHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools,
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case 'outlook_auth_login':
            return await this.handleAuthLogin();

          case 'outlook_auth_status':
            return await this.handleAuthStatus();

          case 'outlook_calendar_list':
            return await this.handleCalendarList();

          case 'outlook_calendar_events_list':
            return await this.handleCalendarEventsList(args);

          case 'outlook_calendar_event_create':
            return await this.handleCalendarEventCreate(args);

          case 'outlook_mail_folders_list':
            return await this.handleMailFoldersList();

          case 'outlook_mail_messages_list':
            return await this.handleMailMessagesList(args);

          case 'outlook_mail_message_get':
            return await this.handleMailMessageGet(args);

          case 'outlook_mail_send':
            return await this.handleMailSend(args);

          default:
            throw new McpError(
              ErrorCode.MethodNotFound,
              `Unknown tool: ${name}`
            );
        }
      } catch (error: any) {
        if (error instanceof McpError) {
          throw error;
        }
        
        if (error.statusCode === 401) {
          throw new McpError(
            ErrorCode.InvalidRequest,
            'Authentication required. Please use outlook_auth_login first.'
          );
        }

        throw new McpError(
          ErrorCode.InternalError,
          `Tool execution failed: ${error.message}`
        );
      }
    });
  }

  private async handleAuthLogin(): Promise<any> {
    return new Promise((resolve, reject) => {
      const app = express();
      let authCodeReceived = false;

      app.get('/auth/callback', async (req, res) => {
        const code = req.query.code as string;

        if (!code) {
          res.send('Error: No authorization code received');
          reject(new Error('No authorization code received'));
          return;
        }

        try {
          const tokenResponse = await this.authManager.acquireTokenByCode(code);
          authCodeReceived = true;

          res.send(`
            <html>
              <body>
                <h2>Authentication Successful!</h2>
                <p>You can close this window and return to Claude.</p>
                <script>window.close();</script>
              </body>
            </html>
          `);

          if (this.authServer) {
            this.authServer.close();
            this.authServer = null;
          }

          resolve({
            content: [
              {
                type: 'text',
                text: `Successfully authenticated as ${tokenResponse.account.username}`,
              },
            ],
          });
        } catch (error: any) {
          res.send('Authentication failed. Check the console for details.');
          reject(error);
        }
      });

      this.authServer = app.listen(config.port, async () => {
        try {
          const authUrl = await this.authManager.getAuthUrl();
          await open(authUrl);

          resolve({
            content: [
              {
                type: 'text',
                text: `Opening browser for authentication. Please complete the login process.`,
              },
            ],
          });
        } catch (error) {
          if (this.authServer) {
            this.authServer.close();
            this.authServer = null;
          }
          reject(error);
        }
      });

      // Timeout after 5 minutes
      setTimeout(() => {
        if (!authCodeReceived && this.authServer) {
          this.authServer.close();
          this.authServer = null;
          reject(new Error('Authentication timeout'));
        }
      }, 300000);
    });
  }

  private async handleAuthStatus() {
    try {
      const user = await this.graphClient.getUser();
      return {
        content: [
          {
            type: 'text',
            text: `Authenticated as: ${user.displayName} (${user.mail || user.userPrincipalName})`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: 'Not authenticated. Please use outlook_auth_login first.',
          },
        ],
      };
    }
  }

  private async handleCalendarList() {
    const calendars = await this.graphClient.listCalendars();
    const calendarList = calendars.value
      .map((cal: any) => `- ${cal.name} (ID: ${cal.id}, Default: ${cal.isDefaultCalendar})`)
      .join('\n');

    return {
      content: [
        {
          type: 'text',
          text: `Found ${calendars.value.length} calendar(s):\n${calendarList}`,
        },
      ],
    };
  }

  private async handleCalendarEventsList(args: any) {
    const events = await this.graphClient.listCalendarEvents(
      args.calendarId,
      args.startDateTime,
      args.endDateTime
    );

    if (events.value.length === 0) {
      return {
        content: [
          {
            type: 'text',
            text: 'No events found in the specified time range.',
          },
        ],
      };
    }

    const eventList = events.value
      .map((event: any) => {
        const start = new Date(event.start.dateTime);
        const end = new Date(event.end.dateTime);
        return `- ${event.subject} (${start.toLocaleString()} - ${end.toLocaleString()})`;
      })
      .join('\n');

    return {
      content: [
        {
          type: 'text',
          text: `Found ${events.value.length} event(s):\n${eventList}`,
        },
      ],
    };
  }

  private async handleCalendarEventCreate(args: any) {
    const eventData: any = {
      subject: args.subject,
      start: {
        dateTime: args.startDateTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: args.endDateTime,
        timeZone: 'UTC',
      },
    };

    if (args.body) {
      eventData.body = {
        contentType: 'Text',
        content: args.body,
      };
    }

    if (args.location) {
      eventData.location = {
        displayName: args.location,
      };
    }

    if (args.attendees && args.attendees.length > 0) {
      eventData.attendees = args.attendees.map((email: string) => ({
        emailAddress: {
          address: email,
        },
        type: 'required',
      }));
    }

    if (args.isOnlineMeeting) {
      eventData.isOnlineMeeting = true;
      eventData.onlineMeetingProvider = 'teamsForBusiness';
    }

    const createdEvent = await this.graphClient.createCalendarEvent(eventData, args.calendarId);

    return {
      content: [
        {
          type: 'text',
          text: `Successfully created event: ${createdEvent.subject} (ID: ${createdEvent.id})`,
        },
      ],
    };
  }

  private async handleMailFoldersList() {
    const folders = await this.graphClient.listMailFolders();
    const folderList = folders.value
      .map((folder: any) => `- ${folder.displayName} (ID: ${folder.id}, Items: ${folder.totalItemCount})`)
      .join('\n');

    return {
      content: [
        {
          type: 'text',
          text: `Found ${folders.value.length} mail folder(s):\n${folderList}`,
        },
      ],
    };
  }

  private async handleMailMessagesList(args: any) {
    const messages = await this.graphClient.listMessages(args.folderId, args.filter);

    if (messages.value.length === 0) {
      return {
        content: [
          {
            type: 'text',
            text: 'No messages found.',
          },
        ],
      };
    }

    const messageList = messages.value
      .slice(0, 10) // Limit to 10 messages for readability
      .map((msg: any) => {
        const from = msg.from?.emailAddress?.address || 'Unknown';
        const date = new Date(msg.receivedDateTime).toLocaleString();
        return `- ${msg.subject} (From: ${from}, Date: ${date})`;
      })
      .join('\n');

    return {
      content: [
        {
          type: 'text',
          text: `Found ${messages.value.length} message(s) (showing first 10):\n${messageList}`,
        },
      ],
    };
  }

  private async handleMailMessageGet(args: any) {
    const message = await this.graphClient.getMessage(args.messageId);
    
    return {
      content: [
        {
          type: 'text',
          text: `Subject: ${message.subject}
From: ${message.from?.emailAddress?.address || 'Unknown'}
To: ${message.toRecipients?.map((r: any) => r.emailAddress.address).join(', ') || 'Unknown'}
Date: ${new Date(message.receivedDateTime).toLocaleString()}

Body:
${message.body?.content || 'No content'}`,
        },
      ],
    };
  }

  private async handleMailSend(args: any) {
    const message: any = {
      subject: args.subject,
      body: {
        contentType: args.isHtml ? 'HTML' : 'Text',
        content: args.body,
      },
      toRecipients: args.to.map((email: string) => ({
        emailAddress: {
          address: email,
        },
      })),
    };

    if (args.cc && args.cc.length > 0) {
      message.ccRecipients = args.cc.map((email: string) => ({
        emailAddress: {
          address: email,
        },
      }));
    }

    if (args.bcc && args.bcc.length > 0) {
      message.bccRecipients = args.bcc.map((email: string) => ({
        emailAddress: {
          address: email,
        },
      }));
    }

    await this.graphClient.sendMail(message);

    return {
      content: [
        {
          type: 'text',
          text: `Successfully sent email: ${args.subject}`,
        },
      ],
    };
  }

  private async cleanup() {
    if (this.authServer) {
      this.authServer.close();
      this.authServer = null;
    }
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Outlook MCP server running');
  }
}

const server = new OutlookMCPServer();
server.run().catch(console.error);