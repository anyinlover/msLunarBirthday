// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import 'isomorphic-fetch';
import {
  DeviceCodeCredential,
  TokenCachePersistenceOptions,
  DeviceCodePromptCallback,
  useIdentityPlugin,
} from '@azure/identity';
import { cachePersistencePlugin } from '@azure/identity-cache-persistence';
import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import {
  User,
  Event,
  Calendar,
  CalendarColor,
  Importance,
  FreeBusyStatus,
} from '@microsoft/microsoft-graph-types';
// prettier-ignore
import { TokenCredentialAuthenticationProvider } from
  '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

import { AppSettings } from './appSettings';

let _settings: AppSettings | undefined = undefined;
let _deviceCodeCredential: DeviceCodeCredential | undefined = undefined;
let _userClient: Client | undefined = undefined;

export function initializeGraphForUserAuth(
  settings: AppSettings,
  deviceCodePrompt: DeviceCodePromptCallback,
) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;
  useIdentityPlugin(cachePersistencePlugin);
  const tokenCachePersistenceOptions: TokenCachePersistenceOptions = {
    enabled: true, // Enable persistent token caching
    name: 'msgraph', // Optional, default cache name, can be customized
    unsafeAllowUnencryptedStorage: true,
  };

  _deviceCodeCredential = new DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    tokenCachePersistenceOptions,
    userPromptCallback: deviceCodePrompt,
  });

  const authProvider = new TokenCredentialAuthenticationProvider(
    _deviceCodeCredential,
    {
      scopes: settings.graphUserScopes,
    },
  );

  _userClient = Client.initWithMiddleware({
    authProvider: authProvider,
  });
}
// </UserAuthConfigSnippet>

export async function getCalendarsAsync(): Promise<PageCollection> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  return _userClient.api('me/calendars').get();
}

export async function getEventsAsync(
  calendarID: string,
): Promise<PageCollection> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  return _userClient.api(`me/calendars/${calendarID}/events`).get();
}

export async function createCalendarAsync(
  calendarName: string,
  calendarColor: CalendarColor,
): Promise<Calendar> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  const calendar: Calendar = {
    name: calendarName,
    color: calendarColor,
  };
  return _userClient.api('me/calendars').post(calendar);
}

export async function createEventAsync(
  calendarID: string,
  subject: string,
  start: string,
  end: string,
  isAllDay: boolean = false,
  isReminderOn: boolean = false,
  reminderMinutesBeforeStart: number = 0,
  importance: Importance = 'normal',
  showAs: FreeBusyStatus = 'busy',
  categories: string[] = [],
): Promise<Event> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  const event: Event = {
    subject: subject,
    start: { dateTime: start, timeZone: 'Asia/Shanghai' },
    end: { dateTime: end, timeZone: 'Asia/Shanghai' },
    isAllDay: isAllDay,
    isReminderOn: isReminderOn,
    reminderMinutesBeforeStart: reminderMinutesBeforeStart,
    showAs: showAs,
    importance: importance,
    categories: categories,
  };
  return _userClient?.api(`me/calendars/${calendarID}/events`).post(event);
}
