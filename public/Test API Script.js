// Make an API request to check calendar availability
fetch('https://graph.microsoft.com/v1.0/me/calendar/getSchedule', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + accessToken
  },
  body: JSON.stringify({
    schedules: ['calendar_id'],
    startTime: {
      dateTime: '2023-09-19T08:00:00',
      timeZone: 'UTC'
    },
    endTime: {
      dateTime: '2023-09-19T17:00:00',
      timeZone: 'UTC'
    }
  })
})
.then(response => response.json())
.then(data => {
  // Process the API response to determine availability
  const isBusy = /* Logic to check availability */;
  
  // Update your website with the availability status
  if (isBusy) {
    document.getElementById('availability').textContent = 'Busy';
  } else {
    document.getElementById('availability').textContent = 'Available';
  }
})
.catch(error => {
  console.error('Error fetching calendar data: ', error);
});
