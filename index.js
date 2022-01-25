const express = require('express')
let axios = require("axios");
require('dotenv').config()
var qs = require('qs');
var cron = require('node-cron');
var bodyParser = require('body-parser')
var mergeRanges = require('merge-ranges');


var jsonParser = bodyParser.json()

var jsonParser = bodyParser.json({ extended: false })

const app = express()
const port = 3000

let token = null;

let generateToken = () => {
    var data = qs.stringify({
        'client_id': process.env.CLIENT_ID,
        'client_secret': process.env.CLIENT_SECRET,
        'grant_type': 'client_credentials',
        'scope': 'https://graph.microsoft.com/.default',
        'redirect_url': 'http://localhost'
    });
    var config = {
        method: 'get',
        url: `https://login.microsoftonline.com/ee5ff5a4-0beb-463d-a383-b4de2845f3ca/oauth2/v2.0/token`,
        headers: {
            'Authorization': 'Basic SFBNYXJ0ZWNoXFZsYWQuUm9taWxhOmtrZnAqUzhRN3lSd1E5Yg==',
            'Content-Type': 'application/x-www-form-urlencoded',
            'Cookie': 'x-ms-gateway-slice=estsfd; stsservicecookie=estsfd; buid=0.ARAApPVf7usLPUajg7TeKEXzyoBgm3LU6ddPjbgQviMoZIwQAAA.AQABAAEAAAD--DLA3VO7QrddgJg7Wevr2HBcQA435h4_pmWRdqJiQHEzjSg5kzG9jU5ICfl4oX5Sn2uIdmmeeg96KMz7kt6w-DALtWg_rDwZNI0KoPK5UXQE5PnNbZpmxJ_NT8gIV64gAA; esctx=AQABAAAAAAD--DLA3VO7QrddgJg7WevroVQcSHkbWa_qtfoIlkTg8tyL6sFFovz2RsZH2c-48AsmhYo2oTgsqejRwxPlKsyTFL05o-sImUYbcxVAR3vIcXyNsdr6eoiwB5AH4clrIC5E1bssNkWDxmImcuuZpGoL2p4DY52oK78F-egzsFFUjk538yZuWmrQ3x2rqnNTmV8gAA; fpc=AjAmdk1acNJDl4_Kmk9qFJc-iZtFAQAAACWKf9kOAAAA'
        },
        data: data
    };

    axios(config)
        .then(function (response) {
            token = response.data.access_token;
        })
        .catch(function (error) {
            console.log(error);
        });
}

generateToken();

cron.schedule('0 */50 * * * *', () => {
    console.log("apelat")
    generateToken();
});

app.get('/', (req, res) => {
    console.log(token);
    res.send('Hello World!')
})

app.post('/getSchedule/', jsonParser, (req, res) => {
    console.log(req.body);
    var data = JSON.stringify({
        "schedules": [req.body.email],
        "startTime": {
            "dateTime": `${req.body.date}T09:00:00`,
            "timeZone": "GTB Standard Time"
        },
        "endTime": {
            "dateTime": `${req.body.date}T18:00:00`,
            "timeZone": "GTB Standard Time"
        },
        "availabilityViewInterval": "15"
    });

    console.log(data);
    var config = {
        method: 'post',
        url: `https://graph.microsoft.com/v1.0/users/${req.body.email}/calendar/getSchedule`,
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            'Prefer': 'outlook.timezone="Europe/Bucharest"'
        },
        data: data
    };

    axios(config)
        .then(function (response) {
            // let busyIntervals = [];
            // console.log(response.data);
            // response.data.value.forEach(value => {
            //     value.scheduleItems.forEach(item => {
            //         if (item.status != "free") {
            //             let startTime = new Date(item.start.dateTime);
            //             let endTime = new Date(item.end.dateTime);
            //             let hasModifiedInterval = false;
            //             busyIntervals.forEach((interval, i) => {
            //                 if (hasModifiedInterval == false)
            //                     if (startTime >= interval[0]) {
            //                         if (endTime >= interval[1]) {
            //                             busyIntervals[i][1] = endTime;
            //                             hasModifiedInterval = true;
            //                         }
            //                     }
            //                     else
            //                         if (endTime <= interval[1]) {
            //                             if (startTime <= interval[0]) {
            //                                 busyIntervals[i][0] = startTime;
            //                                 hasModifiedInterval = true;
            //                             }
            //                         }
            //             })
            //             if (hasModifiedInterval == false) {
            //                 busyIntervals.push([startTime, endTime]);
            //             }
            //         }
            //     })
            // })

            let availableDates = [];
            let startTime = `${req.body.date}T11:00:00`
            let endTime = `${req.body.date}T19:00:00`
            let busyIntervals = [];
            response.data.value.forEach(value => {
                value.scheduleItems.forEach(item => {
                    if (item.status != "free") {
                        busyIntervals.push([new Date(item.start.dateTime), new Date(item.end.dateTime)]);
                    }
                })
            })

            busyIntervals = mergeRanges(busyIntervals);
            console.log(busyIntervals);
            if (busyIntervals.length) {
                while (new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000).getTime() <= new Date(busyIntervals[0][0]).getTime()) {
                    availableDates.push(new Date(startTime).toString());
                    startTime = new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000)
                }
                busyIntervals.forEach((interval, ii) => {
                    startTime = new Date(interval[1]).toString();
                    if (ii == busyIntervals.length - 1) {
                        while (new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000).getTime() <= new Date(endTime).getTime()) {
                            availableDates.push(new Date(startTime).toString());
                            startTime = new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000)
                        }
                    }
                    else {
                        let runSearch = true;
                        while (runSearch) {
                            busyIntervals.forEach(intv => {
                                if (new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000).getTime() >= intv[0] && new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000).getTime() <= intv[1]) {
                                    runSearch = false;
                                }
                            })
                            if (runSearch == true) {
                                availableDates.push(new Date(startTime).toString());
                                startTime = new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000)
                                if (new Date(startTime).getTime() <= new Date(endTime).getTime())
                                    runSearch = false;
                            }
                        }
                    }
                })
            }
            else {
                while (new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000).getTime() <= new Date(endTime).getTime()) {
                    availableDates.push(new Date(startTime).toString());
                    startTime = new Date(new Date(startTime).getTime() + req.body.timeInterval * 60000)
                }
            }

            res.json(availableDates);
        })
        .catch(function (error) {
            console.log(error);
        });
})

app.post('/createEvent', jsonParser, (req, res) => {
    let data = {
        "subject": req.body.subject,
        "body": {
            "contentType": "HTML",
            "content": req.body.content
        },
        "start": {
            "dateTime": "2022-03-15T12:00:00",
            "timeZone": "GTB Standard Time"
        },
        "end": {
            "dateTime": "2022-03-15T12:00:00",
            "timeZone": "GTB Standard Time"
        },
        "attendees": [
            {
                "emailAddress": {
                    "address": "vlad.romila@logdms.com",
                    "name": "Romila Vlad"
                },
                "type": "required"
            },
            {
                "emailAddress": {
                    "address": "romilavlad2001@gmail.com",
                    "name": "Romila Vlad"
                },
                "type": "required"
            }
        ],
        "isOnlineMeeting": true,
        "onlineMeetingProvider": "teamsForBusiness"
    }
    var config = {
        method: 'post',
        url: `https://graph.microsoft.com/v1.0/users/${req.body.email}/calendar/events`,
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            'Prefer': 'outlook.timezone="GTB Standard Time"'
        },
        data: data
    };
    axios(config)
        .then(response => {
            console.log(response.data)
            res.send(response.data);
        })
})

app.listen(process.env.PORT, () => {
    console.log(`Example app listening on port ${process.env.PORT}`)
})