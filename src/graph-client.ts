import { Client } from '@microsoft/microsoft-graph-client';
import { AuthManager } from './auth.js';
import 'isomorphic-fetch';

export class GraphClient {
  private authManager: AuthManager;

  constructor(authManager: AuthManager) {
    this.authManager = authManager;
  }

  private async getAuthenticatedClient(): Promise<Client> {
    const accessToken = await this.authManager.getAccessToken();
    
    if (!accessToken) {
      throw new Error('No access token available. Please authenticate first.');
    }

    return Client.init({
      authProvider: (callback) => {
        callback(null, accessToken);
      },
    });
  }

  async getUser(): Promise<any> {
    const client = await this.getAuthenticatedClient();
    return await client.api('/me').get();
  }

  async listCalendars(): Promise<any> {
    const client = await this.getAuthenticatedClient();
    return await client.api('/me/calendars').get();
  }

  async listCalendarEvents(calendarId?: string, startDateTime?: string, endDateTime?: string): Promise<any> {
    const client = await this.getAuthenticatedClient();
    
    // Default to next 7 days if not specified
    const start = startDateTime || new Date().toISOString();
    const end = endDateTime || new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
    
    // Use the correct endpoint for calendar events
    let endpoint = calendarId 
      ? `/me/calendars/${calendarId}/events`
      : '/me/events'; // Default calendar
    
    let query = client.api(endpoint)
      .filter(`start/dateTime ge '${start}' and end/dateTime le '${end}'`)
      .orderby('start/dateTime')
      .top(50);
    
    return await query.get();
  }

  async createCalendarEvent(eventData: any, calendarId?: string): Promise<any> {
    const client = await this.getAuthenticatedClient();
    
    const endpoint = calendarId 
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';
    
    return await client.api(endpoint).post(eventData);
  }

  async updateCalendarEvent(eventId: string, updates: any): Promise<any> {
    const client = await this.getAuthenticatedClient();
    
    return await client.api(`/me/events/${eventId}`).patch(updates);
  }

  async deleteCalendarEvent(eventId: string): Promise<void> {
    const client = await this.getAuthenticatedClient();
    
    await client.api(`/me/events/${eventId}`).delete();
  }

  async listMailFolders(): Promise<any> {
    const client = await this.getAuthenticatedClient();
    return await client.api('/me/mailFolders').get();
  }

  async listMessages(folderId?: string, filter?: string): Promise<any> {
    const client = await this.getAuthenticatedClient();
    
    let query = folderId 
      ? client.api(`/me/mailFolders/${folderId}/messages`)
      : client.api('/me/messages');
    
    query = query.orderby('receivedDateTime desc').top(50);
    
    if (filter) {
      query = query.filter(filter);
    }
    
    return await query.get();
  }

  async getMessage(messageId: string): Promise<any> {
    const client = await this.getAuthenticatedClient();
    return await client.api(`/me/messages/${messageId}`).get();
  }

  async sendMail(message: any): Promise<void> {
    const client = await this.getAuthenticatedClient();
    
    const sendMail = {
      message: message,
      saveToSentItems: true,
    };
    
    await client.api('/me/sendMail').post(sendMail);
  }

  async replyToMessage(messageId: string, comment: string, replyAll: boolean = false): Promise<void> {
    const client = await this.getAuthenticatedClient();
    
    const endpoint = replyAll ? `/me/messages/${messageId}/replyAll` : `/me/messages/${messageId}/reply`;
    
    await client.api(endpoint).post({
      comment: comment,
    });
  }

  async moveMessage(messageId: string, destinationFolderId: string): Promise<any> {
    const client = await this.getAuthenticatedClient();
    
    return await client.api(`/me/messages/${messageId}/move`).post({
      destinationId: destinationFolderId,
    });
  }
}