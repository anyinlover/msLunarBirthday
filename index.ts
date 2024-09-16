// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// <ProgramSnippet>
import { readFile } from 'fs/promises';
import { DeviceCodeInfo } from '@azure/identity';
import { FreeBusyStatus, ItemBody } from '@microsoft/microsoft-graph-types';
import { LunarDay, SolarDay } from 'tyme4ts';
import { addDays, format } from 'date-fns';

import settings, { AppSettings } from './appSettings';
import * as graphHelper from './graphHelper';

const yearLimit: number = 120;
const birthDayPath: string = './lunarBirthdays.json';
const calendarName = 'LunarBirthdays';
const calendarColor = 'lightRed';

interface BirthdayInfo {
  lunarStr: string;
  solarDate: Date;
  title: string;
}

async function main() {
  // Initialize Graph
  initializeGraph(settings);
  const calendarID: string | undefined = await createBirthdayCalendarAsync();
  const births = await readJsonFile(birthDayPath);
  const birthdays = births
    .map(([name, lunarBirthdayStr]) =>
      calculateBirthday(lunarBirthdayStr, name),
    )
    .flat();
  const tasksPromises = birthdays.map(
    async ({ lunarStr, solarDate, title }) => {
      await createEventAsync(calendarID!, title, solarDate, lunarStr);
    },
  );
  for (let i = 0; i < tasksPromises.length; i += 2) {
    await Promise.all(tasksPromises.slice(i, i + 2));
  }
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
  start: Date,
  content: string,
) {
  const end: Date = addDays(start, 1);
  const isAllDay: boolean = true;
  const isReminderOn: boolean = true;
  const reminderMinutesBeforeStart: number = 0;
  const showAs: FreeBusyStatus = 'free';
  const body: ItemBody = { content, contentType: 'text' };
  try {
    await graphHelper.createEventAsync(
      calendarID,
      subject,
      format(start, "yyyy-MM-dd'T'HH:mm:ss.SSS"),
      format(end, "yyyy-MM-dd'T'HH:mm:ss.SSS"),
      isAllDay,
      isReminderOn,
      reminderMinutesBeforeStart,
      showAs,
      body,
    );
  } catch (err) {
    console.log(`Error create event: ${err}`);
    throw err;
  }
}

async function createBirthdayCalendarAsync() {
  try {
    const calendar = await graphHelper.createCalendarAsync(
      calendarName,
      calendarColor,
    );
    return calendar.id;
  } catch (err) {
    console.log(`Error create calendar: ${err}`);
    throw err;
  }
}

async function readJsonFile(filePath: string): Promise<[string, string][]> {
  const fileContent = await readFile(filePath, 'utf8');
  const birthdays = JSON.parse(fileContent);
  return Object.entries(birthdays);
}

function calculateBirthday(
  lunarBirthdayStr: string,
  name: string,
): BirthdayInfo[] {
  const [birthYear, month, day] = lunarBirthdayStr.split('-').map(Number);
  const currentYear = new Date().getFullYear();
  const range = [];
  for (let i = currentYear; i <= birthYear + yearLimit; i++) {
    range.push(i);
  }
  const lunarDays = range.map((year) => {
    const lunar: LunarDay = LunarDay.fromYmd(year, month, day);

    const lunarStr: string = lunar.toString();
    const solar: SolarDay = lunar.getSolarDay();
    const solarDate: Date = new Date(
      solar.getYear(),
      solar.getMonth() - 1,
      solar.getDay(),
    );
    const age: number = year - birthYear + 1;
    const title: string = `${name}'s ${age}th lunar birthday`;

    return {
      lunarStr,
      solarDate,
      title,
    };
  });

  return lunarDays;
}
