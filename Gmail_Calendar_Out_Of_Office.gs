function activateOutOfOffice() {
  try {
    const calendarId = 'primary';
    const keywords = ['Vacation', 'OOO', 'Out-of-Office', 'Out of office', 'Hols', 'Holiday', 'LW', 'Long Weekend', 'Bank Holiday', 'Christmas', 'AL', 'Annual Leave'];
    const events = findEvents(calendarId, keywords);

    if (events.length === 0) {
      Logger.log('No matching events found');
      return;
    }

    const event = events[0];
    const { start, end } = event;
    const currentTime = new Date();
    const formattedStartDate = formatDate(start.dateTime || start.date);
    const formattedEndDate = formatDate(end.dateTime || end.date);
    const returnDate = new Date(end.dateTime || end.date);
    returnDate.setDate(returnDate.getDate() + 1); // Add one day for return date
    const formattedReturnDate = formatDate(returnDate.toISOString());
    const responseSubject = `Out-of-Office from ${formattedStartDate} to ${formattedEndDate}`;

    Logger.log(`Current time: ${currentTime.toISOString()}`);
    Logger.log(`Event start: ${new Date(start.dateTime || start.date).toISOString()}`);
    Logger.log(`Event end: ${new Date(end.dateTime || end.date).toISOString()}`);

    if (isEventActiveOrUpcoming(currentTime, start, end)) {
      updateVacationSettings(start, end, responseSubject, formattedReturnDate);
      Logger.log(`Out-of-office activated for event: ${event.summary} from ${formattedStartDate} to ${formattedEndDate}`);
    } else {
      Logger.log(`Event not active or upcoming: ${event.summary} from ${formattedStartDate} to ${formattedEndDate}`);
    }
  } catch (error) {
    Logger.log(`Error in activateOutOfOffice: ${error.toString()}`);
  }
}

function isEventActiveOrUpcoming(currentTime, startDate, endDate) {
  const eventStart = new Date(startDate.dateTime || startDate.date);
  const eventEnd = new Date(endDate.dateTime || endDate.date);
  const oneDayBeforeStart = new Date(eventStart.getTime() - 24 * 60 * 60 * 1000);

  return currentTime >= oneDayBeforeStart && currentTime <= eventEnd;
}

function updateVacationSettings(start, end, responseSubject, formattedReturnDate) {
  try {
    const existingSettings = Gmail.Users.Settings.getVacation('me');
    const vacationSettings = {
      enableAutoReply: true,
      responseSubject,
      restrictToContacts: false,
      restrictToDomain: existingSettings.restrictToDomain,
      restrictToUsers: existingSettings.restrictToUsers,
      startTime: Date.parse(start.dateTime || start.date),
      endTime: Date.parse(end.dateTime || end.date)
    };

    // Always generate fresh message with current dates for maximum automation
    vacationSettings.responseBodyHtml = generateDefaultResponseBody(formattedReturnDate);
    Logger.log('Generated fresh out-of-office message with updated return date.');

    //If you sometimes write custom holiday messages but still want date automation, you could implement the below hybrid approach that looks for date placeholders in existing messages:
    
    Gmail.Users.Settings.updateVacation(vacationSettings, 'me');
    Logger.log('Vacation settings updated successfully');
  } catch (error) {
    Logger.log(`Error updating vacation settings: ${error.toString()}`);
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

function formatDate(dateString, timeZone = 'GMT') {
  return Utilities.formatDate(new Date(dateString), timeZone, 'dd MMM yyyy');
}

function isEventActive(currentTime, startDate, endDate) {
  return currentTime >= new Date(startDate.dateTime) && currentTime <= new Date(endDate.dateTime);
}

function generateDefaultResponseBody(formattedReturnDate) {
  return `
  <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6; max-width: 600px;">
    <p style="margin: 0 0 16px 0; font-size: 15px;">Thank you for your email.</p>
    
    <p style="margin: 0 0 16px 0; font-size: 15px;">I'm currently out of the office and will return on <strong>${formattedReturnDate}</strong>. During this time, I'll have limited access to email and may not be able to respond immediately.</p>
    
    <p style="margin: 0 0 20px 0; font-size: 15px;">For urgent matters, please reach out to:</p>
    
    <table style="max-width: 340px; margin: 0 0 24px 0; border-collapse: collapse; background-color: #f8f9fa; border-radius: 8px; overflow: hidden; border: 1px solid #e0e0e0;">
      <tr>
        <td style="padding: 20px; text-align: center; background-color: #0066cc; width: 80px;">
          <div style="font-size: 36px; color: white;">👤</div>
        </td>
        <td style="padding: 20px;">
          <h3 style="margin: 0 0 4px 0; color: #1a1a1a; font-size: 17px; font-weight: 600;">Firstname Surname</h3>
          <p style="margin: 0 0 12px 0; color: #666; font-size: 13px;">Job Role</p>
          <a href="mailto:email@email.com" style="color: #0066cc; text-decoration: none; font-weight: 600; font-size: 14px;">email@email.com →</a>
        </td>
      </tr>
    </table>
    
    <p style="margin: 0; font-size: 14px; color: #666;">Thank you for your understanding.</p>
  </div>`;
}
