'use strict';

const { SessionGraph, canUseBrowser } = require('./graph');
const { TeamsBrowser } = require('./browser');
const { McpOperationError } = require('./errors');

const id = encodeURIComponent;
const literal = text => encodeURIComponent(text.replaceAll("'", "''"));

class Microsoft365Service {
  constructor(runtime, { graph = new SessionGraph(runtime), browser = new TeamsBrowser(runtime) } = {}) {
    this.runtime = runtime;
    this.graph = graph;
    this.browser = browser;
  }

  async withBrowser(backend, apiOperation, browserOperation) {
    if (backend === 'browser') return browserOperation();
    try { return await apiOperation(); }
    catch (error) {
      if (backend === 'graph' || !canUseBrowser(error)) throw error;
      return browserOperation();
    }
  }

  async searchPeople({ query, limit = 20 }) {
    const params = new URLSearchParams({ $search: JSON.stringify(query), $top: String(limit) });
    const data = await this.graph.request(`/me/people?${params}`);
    if (!Array.isArray(data?.value)) throw new McpOperationError('invalid_response', 'Microsoft Graph did not return a people collection.');
    return { source: 'graph', items: data.value.map(person => ({
      userId: person.id, displayName: person.displayName, userPrincipalName: person.userPrincipalName,
      emails: (person.scoredEmailAddresses || []).map(email => email.address).filter(Boolean),
    })) };
  }

  async resolveChat({ userId }) {
    await this.runtime.requireSession();
    const client = await this.graph.getTokenClient();
    const result = await client.resolveConversation(userId);
    if (!result?.success || !result.conversationId) {
      throw new McpOperationError('chat_resolution_failed', 'Could not resolve a verified one-to-one chat with this person. No message was sent.');
    }
    return { source: 'teams_session', chatId: result.conversationId, participantUserId: userId };
  }

  async searchMessages({ query, limit = 20, offset = 0, backend = 'auto' }) {
    return this.withBrowser(backend, async () => {
      const data = await this.graph.request('/search/query', { method: 'POST', readOnly: true, body: {
        requests: [{ entityTypes: ['chatMessage'], query: { queryString: query }, from: offset, size: limit }],
      } });
      const containers = (data?.value || []).flatMap(value => value.hitsContainers || []);
      const moreResultsAvailable = containers.some(container => container.moreResultsAvailable);
      return { source: 'graph', query, items: containers.flatMap(container => container.hits || []),
        moreResultsAvailable, ...(moreResultsAvailable ? { nextOffset: offset + limit } : {}) };
    }, () => {
      if (offset) throw new McpOperationError('unsupported_pagination', 'Browser search returns the visible first page. Graph pagination is unavailable for this session.');
      return this.browser.search('messages', query, limit);
    });
  }

  async searchDirectory(category, { query, limit = 20, backend = 'auto' }) {
    return this.withBrowser(backend, async () => {
      const endpoint = category === 'chats' ? '/me/chats?$expand=members&$top=50' : '/me/joinedTeams';
      const matches = [];
      let cursor;
      let pages = 0;
      const needle = query.toLocaleLowerCase();
      do {
        const result = await this.graph.collection(endpoint, { cursor });
        for (const item of result.items) {
          const haystack = [item.topic, item.displayName, item.description,
            ...(item.members || []).flatMap(member => [member.displayName, member.email])].filter(Boolean).join(' ').toLocaleLowerCase();
          if (haystack.includes(needle)) matches.push(item);
        }
        cursor = result.nextCursor;
        pages++;
      } while (cursor && pages < 10 && matches.length < limit);
      return { source: 'graph', query, items: matches.slice(0, limit), truncated: !!cursor || matches.length > limit,
        coverage: 'Names, descriptions, and available chat member names/emails; scans up to ten API pages. Narrow the query if truncated.' };
    }, () => this.browser.search(category, query, limit));
  }

  async searchOneDrive({ query, limit = 20, backend = 'auto', cursor }) {
    return this.withBrowser(backend,
      () => this.graph.collection(`/me/drive/root/search(q='${literal(query)}')?$top=${limit}`, { cursor }),
      () => {
        if (cursor) throw new McpOperationError('unsupported_pagination', 'The browser cannot continue a Graph cursor. Start a new browser search without a cursor.');
        return this.browser.searchOneDrive(query, limit);
      });
  }

  listOneDrive({ folderId = 'root', limit = 20, cursor }) {
    const folder = folderId === 'root' ? 'root' : `items/${id(folderId)}`;
    return this.graph.collection(`/me/drive/${folder}/children?$top=${limit}`, { cursor });
  }

  listChannels({ teamId, cursor }) {
    return this.graph.collection(`/teams/${id(teamId)}/channels?$select=id,displayName,description,membershipType,webUrl`, { cursor });
  }

  readMessages({ chatId, teamId, channelId, limit = 20, cursor }) {
    const endpoint = chatId ? `/chats/${id(chatId)}/messages` : `/teams/${id(teamId)}/channels/${id(channelId)}/messages`;
    return this.graph.collection(`${endpoint}?$top=${limit}`, { cursor });
  }

  async sendMessage({ chatId, teamId, channelId, replyToId, content }) {
    let endpoint = chatId ? `/chats/${id(chatId)}/messages` : `/teams/${id(teamId)}/channels/${id(channelId)}/messages`;
    if (replyToId) endpoint += `/${id(replyToId)}/replies`;
    const data = await this.graph.request(endpoint, { method: 'POST', body: { body: { contentType: 'text', content } } });
    if (!data?.id) throw new McpOperationError('outcome_unknown', 'Microsoft did not return a message ID. Check Teams before retrying the send.');
    return { source: 'graph', sent: true, message: data };
  }

  listCalendarEvents({ startDateTime, endDateTime, calendarId, limit = 20, cursor }) {
    const calendar = calendarId ? `/me/calendars/${id(calendarId)}` : '/me/calendar';
    const params = new URLSearchParams({ startDateTime, endDateTime, $top: limit, $orderby: 'start/dateTime' });
    return this.graph.collection(`${calendar}/calendarView?${params}`, { cursor });
  }

  async createCalendarEvent({ calendarId, event }) {
    const calendar = calendarId ? `/me/calendars/${id(calendarId)}` : '/me/calendar';
    return { source: 'graph', event: await this.graph.request(`${calendar}/events`, { method: 'POST', body: event }) };
  }

  async updateCalendarEvent({ eventId, updates }) {
    return { source: 'graph', event: await this.graph.request(`/me/events/${id(eventId)}`, { method: 'PATCH', body: updates }) };
  }
}

module.exports = { Microsoft365Service };
