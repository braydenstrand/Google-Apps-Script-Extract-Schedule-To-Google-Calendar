function myFunction() {

  // Scan inbox for the proper message
  const message = getEmailMessage();

  // Extract the starting date from the message
  const startingDate = getStartingDate(message);

  // Extract the schedule times from the message
  const scheduleTimes = extractScheduleTimes(message);

  // Create the schedule days using the starting date and schedule times
  const scheduleDays = createScheduleDays(scheduleTimes, startingDate);

  // Optionally log the created scheduleDays
  //logScheduleDays(scheduleDays);

  // Create the calendar events
  createCalendarEvents(scheduleDays);

}

function getEmailMessage(){
  // Retrieve email message from boss

  const threads = GmailApp.search('from: Boss@Work.com');
  
  let message = threads[0].getMessages()[0];

  // Check if either of the first two message subject lines contain the word schedule. Error if not.
  if (message.getSubject().toLowerCase().match('schedule')) {
    return message;
  } else {
    message = threads[1].getMessages()[0];
    if (message.getSubject().toLowerCase().match('schedule')){
      return message;
    } else {
        throw new Error('First 2 message subjects do not contain the word Schedule');
      }
  }
}

function getStartingDate(message){
  // Extract the year, month, and day for the first day of the schedule
  
  let messagePlainBody = message.getPlainBody();

  // Find the current year, which is followed by 'SUN' in the email message
  let currentYear;
  let match = messagePlainBody.match(/\b(\d{4})\s*SUN\b/);
  if (match){
    currentYear = match[1];
  }else{
    throw new Error('No year found');
  }

  // Find the current month, which is followed by the first day in the email message
  let currentMonth;
  let firstDay;
  match = messagePlainBody.match(/\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*(.{0,8})\s*(\d{1,2})\b/);

  // Define the month mapping
    const monthMap = {
      JAN: 0,
      FEB: 1,
      MAR: 2,
      APR: 3,
      MAY: 4,
      JUN: 5,
      JUL: 6,
      AUG: 7,
      SEP: 8,
      OCT: 9,
      NOV: 10,
      DEC: 11
  };
  if (match){
    currentMonth = monthMap[match[1]];
    firstDay = match[3];
  }else{
    throw new Error('No month/day match');
  }

  let date = [+currentYear, +currentMonth, +firstDay];
  return date;
}

function extractScheduleTimes(message)  {
  // Function to extract the schedule times from the email message body
  // Example return: [ '3:30p',
  // '12a',
  // '3:30p',
  // '12a',
  // 'OFF',
  // 'OFF',
  // 'STAT',
  // 'STAT',
  // '7a',
  // '3:30p' ]

  if (!message || !message.getId) {
    throw new Error('Invalid Gmail message object');
  }

  const messageContent = message.getPlainBody();

  // Extract the proper schedule section
  const employeeNameScheduleStart = messageContent.indexOf("employeeName");
  const employeeNameScheduleEnd = messageContent.indexOf("\nFreddy", employeeNameScheduleStart);
  const employeeNameSchedule = messageContent.substring(employeeNameScheduleStart, employeeNameScheduleEnd).trim();

  // Extract times from the schedule section
  const timePattern = /\b\d{1,2}:\d{2}[ap]?|\b\d{1,2}[ap]|\bOFF\b|\bSTAT\b/gi;
  const employeeNameTimes = employeeNameSchedule.match(timePattern);
  
  // Log schedule times
  console.log('employeeName\'s Schedule Times:', employeeNameTimes);

  return employeeNameTimes;
}

function scheduleDay(year, month, day, startTime, endTime)  {
  // Object that contains 2 Date objects for the start and end of the shift

  this.beginning = parseTimeStringToDate(startTime);
  this.beginning.setFullYear(year, month, day);
  
  this.end = parseTimeStringToDate(endTime);
  // if the day ends with midnight, add one to the day so that the date is correct
  if(this.end.getHours() == 0){
    this.end.setFullYear(year, month, day + 1);
  }else{
    this.end.setFullYear(year, month, day);
  }
}

function parseTimeStringToDate(timeString) {
  // Take the input timeString and return a Date object

  if (!timeString){
    Logger.log('error');
  }
  // Create a new Date object for today
  var today = new Date();
  
  // Extract hours and minutes from the time string
  
  var match = timeString.match(/(\d{1,2}):?(\d{2})?([ap]?)/i);
  if (!match) {
    throw new Error("Invalid time format: " + timeString);
  }
  
  var hours = parseInt(match[1], 10);
  var minutes = match[2] ? parseInt(match[2], 10) : 0;
  var period = match[3] ? match[3].toLowerCase() : '';

  // Adjust hours based on AM/PM
  if (period === 'p' && hours < 12) {
    hours += 12;
  } else if (period === 'a' && hours === 12) {
    hours = 0;
  }

  // Set the hours and minutes in the Date object
  today.setHours(hours);
  today.setMinutes(minutes);
  today.setSeconds(0);
  today.setMilliseconds(0);
  
  return today;
}

function createScheduleDays(scheduleTimes, startingDate){
  // Function to create scheduleDay object for each pair of times in scheduleTimes 
  // If 'off' or 'stat' is found, that spot in the return array will be null

  const scheduleDays = new Array(7);

  for(let i = 0, j = 0; i < scheduleDays.length; i++, j+=2)  {

    // If the current schedule time is 'off' or 'stat', that spot in the array will be null
    timeTemp = scheduleTimes[j].toLowerCase();
    if(timeTemp == 'off' || timeTemp == 'stat'){
      j--;
      continue
    }
    // Create the scheduleDay object
    else{
      scheduleDays[i] = new scheduleDay(startingDate[0], startingDate[1], startingDate[2] + i, scheduleTimes[j], scheduleTimes[j + 1]);
    }
  }

  return scheduleDays;
}

function logScheduleDays(scheduleDays){
  for(let i = 0; i < scheduleDays.length; i++){
     if(scheduleDays[i] == null){
       Logger.log(i + ' OFF');
     }
     else{
      Logger.log(scheduleDays[i].beginning);
      Logger.log(scheduleDays[i].end);
     }
    // if(scheduleDays[i] != null){
    //   Logger.log(i +' test 2');
    // }
  }
}

function createCalendarEvents(scheduleDays){
  // Create the google calendar events

  // Get the calendar
  let myCalendar = CalendarApp.getDefaultCalendar();
  
  var createdEvents = [];

  // Check if the first event has already been created
  for (i = 0; i < scheduleDays.length; i++){
    if (scheduleDays[i] != null){
      const eventsTemp = myCalendar.getEvents(scheduleDays[i].beginning, scheduleDays[i].end);
      for (j = 0; j < eventsTemp.length; j++){
        if(eventsTemp.length > 0 &&  eventsTemp[i].getTitle() == 'Work'){
          throw new Error('First event already exists. Stopping program.');
        }
      }
    }
  }
  
  // Loop through the scheduleDays array
  for (var i = 0; i < scheduleDays.length; i++) {
    // If the index is null, skip it
    if (scheduleDays[i] == null) {
      continue;
    }
    
    // Try to create a google calendar event with reminders
    try {
      createdEvents[i] = myCalendar.createEvent('Work', scheduleDays[i].beginning, scheduleDays[i].end);
      createdEvents[i].removeAllReminders();
      createdEvents[i].addPopupReminder(120);
      createdEvents[i].addPopupReminder(720);
      createdEvents[i].addPopupReminder(1440);
      Logger.log('Created Event: Start: ' + scheduleDays[i].beginning + ' End: ' + scheduleDays[i].end);
    } catch (e) {
      // if error, delete events that were already created so there is not a partial schedule
      deleteCreatedCalendarEvents(createdEvents);
      throw new Error('Error creating event: ' + e.message);
      }
  }
}

function deleteCreatedCalendarEvents(createdEvents){
  createdEvents.forEach(function(event) {
        if (event) {
          event.deleteEvent();
          Logger.log('Event deleted: ' + event.getTitle());
        }
    });
}
