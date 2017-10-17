// server.js
// where your node app starts

// init project
var express = require('express');
var app = express();
var ew = require("ews-javascript-api");
var moment = require("moment");
var events = require('events');
// var eventEmitter = new events.EventEmitter();

app.use(express.static('public'));

ew.EwsLogging.DebugLogEnabled = false;

let service = new ew.ExchangeService(ew.ExchangeVersion.Exchange2013);
service.Credentials = new ew.ExchangeCredentials(process.env.UN, process.env.PW);
service.Url = new ew.Uri("https://outlook.office365.com/Ews/Exchange.asmx");


function addCategory(email, todaysDate, tomorrowsDate, arrivalDate) {
  console.log("adding category!");
  if(arrivalDate == todaysDate && !email.Categories.items.includes('Arrives Today')) {
    console.log("dates match!");
    email.Categories.Add('Arrives Today');
    email.Flag.FlagStatus = ew.ItemFlagStatus.Flagged;
    email.Update();
  } else if(arrivalDate == tomorrowsDate && !email.Categories.items.includes('Arrives Tomorrow'))
  {
    console.log("that's tomorrow!");
    email.Categories.Add('Arrives Tomorrow');
    email.Flag.FlagStatus = ew.ItemFlagStatus.Flagged;
    email.Update();
  }

}

function getArrivalDate(email) {
    console.log("getting arrival date!");
    let subject = email.Subject;
    let body = email.Body.text;
    let todaysDateTime = new Date();
    let todaysDate = moment(todaysDateTime).tz('America/New_York').format('MM/DD/YYYY');
    let tomorrowsDate = moment(todaysDateTime).tz('America/New_York').add(1, 'days').format('MM/DD/YYYY');
    let arrivalDateString = '';
    let arrivalDateTime;
    let arrivalDate;

    if(subject.includes("Booking.com")) {
    //  console.log(body);
      arrivalDateString = body.substring(body.indexOf("Arrival Date .....: ") + 20,body.indexOf(" Departure Date ...: "));
      console.log(arrivalDateString);
      arrivalDateTime = new Date(arrivalDateString);
 
    } else if(subject.includes("[TheBookingButton]")) {
      arrivalDateString = body.substring(body.indexOf("Check In Date: ") + 15,body.indexOf("Check Out Date: "));
      arrivalDateTime = new Date(arrivalDateString);
      // console.log(arrivalDateTime);

    }

    arrivalDate = moment(arrivalDateTime).format('MM/DD/YYYY');
    return {
      "todaysDate": todaysDate,
      "tomorrowsDate": tomorrowsDate,
      "arrivalDate": arrivalDate
    };
}


var categorizeEmail = (itemIDString) => {
  // let emailService = new ew.ExchangeService(ew.ExchangeVersion.Exchange2013);
  // emailService.Credentials = new ew.ExchangeCredentials(process.env.UN, process.env.PW);
  // emailService.Url = new ew.Uri("https://outlook.office365.com/Ews/Exchange.asmx");
  let itemID = new ew.ItemId(itemIDString);
  ew.EmailMessage.Bind(service, itemID).then((response) => {
    console.log(response.Subject);
    let dates = getArrivalDate(response);
    addCategory(response, dates.todaysDate, dates.tomorrowsDate, dates.arrivalDate);
  }, function(error) {
    console.log(error);
  });

}

// get shared box
var sharedAddress = new ew.Mailbox("westside2@ymcanyc.org");
var sharedFolder = new ew.FolderId(ew.WellKnownFolderName.Inbox, sharedAddress);

service.SubscribeToStreamingNotifications(
    [new ew.FolderId(ew.WellKnownFolderName.Inbox)],
    // [sharedFolder],
    ew.EventType.NewMail).then((streamingSubscription) => {
        // console.log(streamingSubscription);
        // Create a streaming connection to the service object, over which events are returned to the client.
        // Keep the streaming connection open for 30 minutes.
        let connection = new ew.StreamingSubscriptionConnection(service, 30);
        connection.AddSubscription(streamingSubscription);
        connection.OnNotificationEvent.push((o, a) => {
          console.log("notification received") //this gives you each notification.
          // ew.EwsLogging.Log(a, true, true);
          let notifications = a.Events;
          // console.log(notifications);
          for(var i = 0; i < notifications.length; i++) {
            // console.log(notifications[i]);
            let itemEvent = notifications[i];
            // console.log(itemEvent);
            if(itemEvent.eventType == ew.EventType.NewMail) {
              let itemId = itemEvent.itemId.UniqueId;
              // eventEmitter.emit('categorize', itemId);
              categorizeEmail(itemId);
            }
          }
        });

        connection.OnDisconnect.push((connection, subscriptionErrorEventArgsInstance) => {
           console.log("disconnected...");
            console.log(subscriptionErrorEventArgsInstance);
           connection.Open();
        });

        connection.Open();

    }, (err) => {
        debugger;
        console.log(err);
});

app.get("/wake", function (req, res) {
  let timeCheck = moment().tz('America/New_York');
  console.log(timeCheck);
  res.sendStatus(200);
});

// listen for requests :)
var listener = app.listen(process.env.PORT, function () {
  console.log('Your app is listening on port ' + listener.address().port);
});
