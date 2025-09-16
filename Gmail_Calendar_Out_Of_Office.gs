function activateOutOfOffice() {
  try {
    const calendarId = 'primary';
    const keywords = ['Vacation', 'OOO', 'Holiday', 'Long Weekend', 'Bank Holiday', 'Christmas', 'Out-of-Office', 'Out of office'];
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

function updateVacationSettings(start, end, responseSubject, formattedEndDate) {
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
    
    /*if (existingSettings.responseBodyHtml && existingSettings.responseBodyHtml.trim() !== '<div dir="ltr"></div>') {
      // Replace any date placeholders in existing message
      vacationSettings.responseBodyHtml = existingSettings.responseBodyHtml.replace(/\{RETURN_DATE\}/g, formattedReturnDate);
      Logger.log('Updated existing message with current return date.');
    } else {
      vacationSettings.responseBodyHtml = generateDefaultResponseBody(formattedReturnDate);
      Logger.log('Generated fresh out-of-office message with updated return date.');
    }*/

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
  return `Hi,<br><br>Thank you for your email,<br><br>I am currently out of the office and will not be available to respond to emails until ${formattedReturnDate}. During this period, my access to email will be limited, and there may be a delay in my response.<br><br>If you require immediate assistance, please contact:<br><br>
  <div style="max-width: 300px; margin: 20px 0; padding: 20px; border: 2px solid #0066cc; border-radius: 10px; background-color: #f8f9fa; font-family: Arial, sans-serif; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
    <div style="text-align: center; margin-bottom: 15px;">
      <div style="width: 60px; height: 60px; background-color: #0066cc; border-radius: 50%; margin: 0 auto 10px; display: flex; align-items: center; justify-content: center; color: white; font-size: 24px; font-weight: bold;">JB</div>
      <h3 style="margin: 0; color: #333; font-size: 18px;">Joe Bloggs</h3>
      <p style="margin: 5px 0 0 0; color: #666; font-size: 14px; font-style: italic;">Senior Manager</p>
    </div>
    <div style="text-align: center;">
      <a href="mailto:joe.bloggs@email.com" style="display: inline-block; padding: 10px 20px; background-color: #0066cc; color: white; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px; transition: background-color 0.3s;">📧 Contact Joe</a>
    </div>
  </div><br>I appreciate your understanding and cooperation during my absence.<br>Thank you for your patience.<br><br>Best Regards.`;
}
