# msLunarBirthday
A tool to build a Chinese lunar birthday calendar in Outlook

You need a `lunarBirthdays.json` file with the following format:

```json
{
    "Me": "1982-01-11",
    "Wife": "1982-07-02",
    "Mother": "1961-04-27"
}
```

Then run `npx ts-node index.ts`, it will create a calendar in Outlook with the birthdays of your family members.