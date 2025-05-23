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
import { AccountResponse } from './types.js';
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
    description: 'Authenticate with Microsoft Outlook accounts. Opens browser windows for OAuth2 authentication. Use this when asked to login to email, calendar, or Outlook.',
    inputSchema: {
      type: 'object',
      properties: {
        account: {
          type: 'string',
          description: 'Optional: Specific account name to authenticate. If not provided, authenticates all configured accounts.',
        },
      },
    },
  },
  {
    name: 'outlook_auth_status',
    description: 'Check authentication status for all configured accounts. Shows which accounts are logged in.',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'calendar_list',
    description: 'List all calendars/agendas from Outlook accounts. Use when asked about calendars, agendas, or schedules.',
    inputSchema: {
      type: 'object',
      properties: {
        account: {
          type: 'string',
          description: 'Optional: Specific account name. If not provided, lists calendars from all accounts.',
        },
      },
    },
  },
  {
    name: 'calendar_events_list',
    description: 'List calendar events/appointments within a date range. Use when asked about meetings, appointments, events, or what\'s in the calendar/agenda.',
    inputSchema: {
      type: 'object',
      properties: {
        account: {
          type: 'string',
          description: 'Optional: Specific account name. If not provided, lists events from all accounts.',
        },
        calendarId: {
          type: 'string',
          description: 'Optional: Calendar ID. Defaults to primary calendar.',
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
    name: 'calendar_event_create',
    description: 'Create a new calendar event/appointment. Use when asked to schedule, create, or add something to the calendar/agenda.',
    inputSchema: {
      type: 'object',
      properties: {
        account: {
          type: 'string',
          description: 'Optional: Specific account name. If not provided, creates event in all accounts.',
        },
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
    name: 'mail_folders_list',
    description: 'List all email/mail folders from Outlook accounts. Use when asked about email folders or mailboxes.',
    inputSchema: {
      type: 'object',
      properties: {
        account: {
          type: 'string',
          description: 'Optional: Specific account name. If not provided, lists folders from all accounts.',
        },
      },
    },
  },
  {
    name: 'mail_messages_list',
    description: 'List email messages. Use when asked to check, read, or show emails/mail.',
    inputSchema: {
      type: 'object',
      properties: {
        account: {
          type: 'string',
          description: 'Optional: Specific account name. If not provided, lists messages from all accounts.',
        },
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
    name: 'mail_message_get',
    description: 'Get a specific email message by ID. Use when asked to read a specific email in detail.',
    inputSchema: {
      type: 'object',
      properties: {
        account: {
          type: 'string',
          description: 'Account name where the message is located',
        },
        messageId: {
          type: 'string',
          description: 'Message ID',
        },
      },
      required: ['account', 'messageId'],
    },
  },
  {
    name: 'mail_send',
    description: 'Send an email message. Use when asked to send, compose, or write an email.',
    inputSchema: {
      type: 'object',
      properties: {
        account: {
          type: 'string',
          description: 'Optional: Specific account to send from. If not provided, sends from the first configured account.',
        },
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
  private authServer: any = null;

  constructor() {
    this.server = new Server(
      {
        name: 'mcp-outlook-server',
        version: '2.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.authManager = new AuthManager();

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
            return await this.handleAuthLogin(args);

          case 'outlook_auth_status':
            return await this.handleAuthStatus();

          case 'calendar_list':
            return await this.handleCalendarList(args);

          case 'calendar_events_list':
            return await this.handleCalendarEventsList(args);

          case 'calendar_event_create':
            return await this.handleCalendarEventCreate(args);

          case 'mail_folders_list':
            return await this.handleMailFoldersList(args);

          case 'mail_messages_list':
            return await this.handleMailMessagesList(args);

          case 'mail_message_get':
            return await this.handleMailMessageGet(args);

          case 'mail_send':
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

  private async executeForAccounts<T>(
    accountName: string | undefined,
    operation: (account: string) => Promise<T>
  ): Promise<AccountResponse<T>[]> {
    const accounts = accountName 
      ? [accountName] 
      : this.authManager.getAllAccountNames();
    
    const results = await Promise.allSettled(
      accounts.map(async (account) => {
        try {
          const data = await operation(account);
          return { account, data };
        } catch (error: any) {
          return { account, data: null as any, error: error.message };
        }
      })
    );

    return results.map((result) => {
      if (result.status === 'fulfilled') {
        return result.value;
      } else {
        return { account: 'unknown', data: null as any, error: result.reason };
      }
    });
  }

  private formatMultiAccountResponse(results: AccountResponse<any>[]): any {
    if (results.length === 1) {
      // Single account response
      const result = results[0];
      if (result.error) {
        throw new Error(`${result.account}: ${result.error}`);
      }
      return result.data;
    } else {
      // Multi-account response
      const successResults = results.filter(r => !r.error);
      const errorResults = results.filter(r => r.error);
      
      let response = '';
      
      if (successResults.length > 0) {
        response = successResults.map(r => `[${r.account}]\n${this.formatSingleResponse(r.data)}`).join('\n\n');
      }
      
      if (errorResults.length > 0) {
        const errors = errorResults.map(r => `[${r.account}] Error: ${r.error}`).join('\n');
        response += (response ? '\n\n' : '') + errors;
      }
      
      return {
        content: [
          {
            type: 'text',
            text: response,
          },
        ],
      };
    }
  }

  private formatSingleResponse(data: any): string {
    if (typeof data === 'string') {
      return data;
    } else if (data && data.content && data.content[0] && data.content[0].text) {
      return data.content[0].text;
    } else {
      return JSON.stringify(data, null, 2);
    }
  }

  private async handleAuthLogin(args: any): Promise<any> {
    const accounts = args.account 
      ? [args.account] 
      : this.authManager.getAllAccountNames();

    if (accounts.length === 1) {
      // Single account login
      return await this.authenticateSingleAccount(accounts[0]);
    } else {
      // Multi-account login
      return await this.authenticateMultipleAccounts(accounts);
    }
  }

  private async authenticateSingleAccount(accountName: string): Promise<any> {
    return new Promise((resolve, reject) => {
      const app = express();
      let authCodeReceived = false;

      app.get('/auth/callback', async (req, res) => {
        const code = req.query.code as string;
        const state = req.query.state as string;

        if (!code) {
          res.send('Error: No authorization code received');
          reject(new Error('No authorization code received'));
          return;
        }

        try {
          const tokenResponse = await this.authManager.acquireTokenByCode(code, state || accountName);
          authCodeReceived = true;

          res.send(`
            <html>
              <body>
                <h2>Authentication Successful!</h2>
                <p>Account: ${state || accountName}</p>
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
                text: `Successfully authenticated ${state || accountName} as ${tokenResponse.account.username}`,
              },
            ],
          });
        } catch (error: any) {
          res.send('Authentication failed. Check the console for details.');
          reject(error);
        }
      });

      this.authServer = app.listen(config.server.port, async () => {
        try {
          const authUrl = await this.authManager.getAuthUrl(accountName);
          await open(authUrl);

          resolve({
            content: [
              {
                type: 'text',
                text: `Opening browser for ${accountName} authentication. Please complete the login process.`,
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

      setTimeout(() => {
        if (!authCodeReceived && this.authServer) {
          this.authServer.close();
          this.authServer = null;
          reject(new Error('Authentication timeout'));
        }
      }, 300000);
    });
  }

  private async authenticateMultipleAccounts(accounts: string[]): Promise<any> {
    let response = 'Starting authentication for multiple accounts...\n\n';
    
    for (const account of accounts) {
      try {
        const result = await this.authenticateSingleAccount(account);
        response += `✅ ${account}: ${this.formatSingleResponse(result)}\n`;
      } catch (error: any) {
        response += `❌ ${account}: ${error.message}\n`;
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: response,
        },
      ],
    };
  }

  private async handleAuthStatus(): Promise<any> {
    const results = await this.executeForAccounts(undefined, async (account) => {
      const graphClient = new GraphClient(this.authManager, account);
      try {
        const user = await graphClient.getUser();
        return `Authenticated as: ${user.displayName} (${user.mail || user.userPrincipalName})`;
      } catch (error) {
        return `Not authenticated`;
      }
    });

    return this.formatMultiAccountResponse(results);
  }

  private async handleCalendarList(args: any): Promise<any> {
    const results = await this.executeForAccounts(args.account, async (account) => {
      const graphClient = new GraphClient(this.authManager, account);
      const calendars = await graphClient.listCalendars();
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
    });

    return this.formatMultiAccountResponse(results);
  }

  private async handleCalendarEventsList(args: any): Promise<any> {
    const results = await this.executeForAccounts(args.account, async (account) => {
      const graphClient = new GraphClient(this.authManager, account);
      const events = await graphClient.listCalendarEvents(
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
    });

    return this.formatMultiAccountResponse(results);
  }

  private async handleCalendarEventCreate(args: any): Promise<any> {
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

    const results = await this.executeForAccounts(args.account, async (account) => {
      const graphClient = new GraphClient(this.authManager, account);
      const createdEvent = await graphClient.createCalendarEvent(eventData, args.calendarId);

      return {
        content: [
          {
            type: 'text',
            text: `Successfully created event: ${createdEvent.subject} (ID: ${createdEvent.id})`,
          },
        ],
      };
    });

    return this.formatMultiAccountResponse(results);
  }

  private async handleMailFoldersList(args: any): Promise<any> {
    const results = await this.executeForAccounts(args.account, async (account) => {
      const graphClient = new GraphClient(this.authManager, account);
      const folders = await graphClient.listMailFolders();
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
    });

    return this.formatMultiAccountResponse(results);
  }

  private async handleMailMessagesList(args: any): Promise<any> {
    const results = await this.executeForAccounts(args.account, async (account) => {
      const graphClient = new GraphClient(this.authManager, account);
      const messages = await graphClient.listMessages(args.folderId, args.filter);

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
        .slice(0, 10)
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
    });

    return this.formatMultiAccountResponse(results);
  }

  private async handleMailMessageGet(args: any): Promise<any> {
    if (!args.account) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        'Account name is required for getting a specific message'
      );
    }

    const graphClient = new GraphClient(this.authManager, args.account);
    const message = await graphClient.getMessage(args.messageId);
    
    return {
      content: [
        {
          type: 'text',
          text: `[${args.account}]
Subject: ${message.subject}
From: ${message.from?.emailAddress?.address || 'Unknown'}
To: ${message.toRecipients?.map((r: any) => r.emailAddress.address).join(', ') || 'Unknown'}
Date: ${new Date(message.receivedDateTime).toLocaleString()}

Body:
${message.body?.content || 'No content'}`,
        },
      ],
    };
  }

  private async handleMailSend(args: any): Promise<any> {
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

    const accountToUse = args.account || this.authManager.getAllAccountNames()[0];
    const graphClient = new GraphClient(this.authManager, accountToUse);
    await graphClient.sendMail(message);

    return {
      content: [
        {
          type: 'text',
          text: `Successfully sent email from ${accountToUse}: ${args.subject}`,
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
    console.error('Outlook MCP server running with multi-account support');
  }
}

const server = new OutlookMCPServer();
server.run().catch(console.error);