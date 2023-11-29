function activateOutOfOffice() {
  const calendarId = 'primary';
  const keywords = ['Out of Office', 'OOO', 'Out-of-Office', 'Vacation', 'Holiday', 'Long Weekend', 'Bank Holiday', 'Christmas', 'Xmas'];
  const events = findEvents(calendarId, keywords);

  if (events.length > 0) {
    const event = events[0];
    const { start, end } = event;
    const currentTime = new Date();
    const formattedStartDate = formatDate(start.dateTime);
    const formattedEndDate = formatDate(end.dateTime);
    const responseSubject = `Out-of-Office from ${formattedStartDate} to ${formattedEndDate}`;

    if (isEventActive(currentTime, start, end)) {
      const existingSettings = Gmail.Users.Settings.getVacation('me');

      if (existingSettings.responseBodyHtml.trim() !== '<div dir="ltr"></div>') {
        Gmail.Users.Settings.updateVacation({
          enableAutoReply: true,
          responseSubject,
          responseBodyHtml: existingSettings.responseBodyHtml,
          restrictToContacts: false,
          restrictToDomain: existingSettings.restrictToDomain,
          restrictToUsers: existingSettings.restrictToUsers,
          startTime: Date.parse(start.dateTime),
          endTime: Date.parse(end.dateTime)
        }, 'me');
      } else {
        const newResponseBody = generateDefaultResponseBody(formattedEndDate);
        Gmail.Users.Settings.updateVacation({
          enableAutoReply: true,
          responseSubject,
          responseBodyHtml: newResponseBody,
          restrictToContacts: false,
          restrictToDomain: existingSettings.restrictToDomain,
          restrictToUsers: existingSettings.restrictToUsers,
          startTime: Date.parse(start.dateTime),
          endTime: Date.parse(end.dateTime)
        }, 'me');
      }
    }
  }
}


function findEvents(calendarId, keywords) {
  for (const keyword of keywords) {
    const eventsForKeyword = Calendar.Events.list(calendarId, {
      timeMin: new Date().toISOString(),
      maxResults: 1,
      singleEvents: true,
      orderBy: 'startTime',
      q: keyword
    }).items;

    if (eventsForKeyword.length > 0) {
      return eventsForKeyword;
    }
  }

  return [];
}

function formatDate(dateString) {
  return Utilities.formatDate(new Date(dateString), 'GMT', 'dd MMM yyyy');
}

function isEventActive(currentTime, startDate, endDate) {
  return currentTime >= new Date(startDate.dateTime) && currentTime <= new Date(endDate.dateTime);
}

function generateDefaultResponseBody(formattedEndDate) {
  return `Hi,<br><br>Thank you for your email,<br><br>I am currently out of the Out-Of-Office and will not be available to respond to emails until ${formattedEndDate}. During this period, my access to email will be limited, and there may be a delay in my response. If you require immediate assistance, please send an email to <a href="mailto:username@email.com">username@email.com</a> in my absence.<br><br>I appreciate your understanding and cooperation during my absence.<br>Thank you for your patience.<br><br>Best Regards.`;
}