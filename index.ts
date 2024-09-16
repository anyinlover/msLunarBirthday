// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// <ProgramSnippet>
import { DeviceCodeInfo } from '@azure/identity';
import {
  Calendar,
  Event,
  Importance,
} from '@microsoft/microsoft-graph-types';

import settings, { AppSettings } from './appSettings';
import * as graphHelper from './graphHelper';

let calendarID: string | undefined = undefined;

async function main() {
  // Initialize Graph
  initializeGraph(settings);
  await createBirthdayCalendarAsync();
}

main();
// </ProgramSnippet>

// <InitializeGraphSnippet>
function initializeGraph(settings: AppSettings) {
  graphHelper.initializeGraphForUserAuth(settings, (info: DeviceCodeInfo) => {
    // Display the device code message to
    // the user. This tells them
    // where to go to sign in and provides the
    // code to use.
    console.log(info.message);
  });
}
// </InitializeGraphSnippet>

async function createEventAsync(
  calendarID: string,
  subject: string,
  start: string,
  end: string,
  importance: Importance,
) {
  try {
    const newEvent = await graphHelper.createEventAsync(
      subject,
      start,
      end,
      importance,
    );
    console.log(`Event created with id ${newEvent.id}`);
  } catch (err) {
    console.log(`Error create event: ${err}`);
  }
}

async function createBirthdayCalendarAsync() {
  try {
    const calendarName = 'LunarBirthday';
    const calendarColor = 'lightRed';
    const calendar = await graphHelper.createCalendarAsync(
      calendarName,
      calendarColor,
    );
    calendarID = calendar.id;
  } catch (err) {
    console.log(`Error create calendar: ${err}`);
  }
}
