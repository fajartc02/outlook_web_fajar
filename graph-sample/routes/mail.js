// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const router = require("express-promise-router").default();
const graph = require("../graph.js");
const dateFns = require("date-fns");
const zonedTimeToUtc = require("date-fns-tz/zonedTimeToUtc");
const iana = require("windows-iana");
const { body, validationResult } = require("express-validator");
const validator = require("validator");

router.get("/", async function (req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect("/");
  } else {
    const params = {
      // active: { calendar: true },
      active: { mail: true },
    };

    // Get the user
    const user = req.app.locals.users[req.session.userId];
    // Convert user's Windows time zone ("Pacific Standard Time")
    // to IANA format ("America/Los_Angeles")
    const timeZoneId = iana.findIana(user.timeZone)[0];
    console.log(`Time zone: ${timeZoneId.valueOf()}`);

    // Calculate the start and end of the current week
    // Get midnight on the start of the current week in the user's timezone,
    // but in UTC. For example, for Pacific Standard Time, the time value would be
    // 07:00:00Z
    var weekStart = zonedTimeToUtc(
      dateFns.startOfWeek(new Date()),
      timeZoneId.valueOf()
    );
    var weekEnd = dateFns.addDays(weekStart, 7);
    console.log(`Start: ${dateFns.formatISO(weekStart)}`);

    try {
      // Get the events

      const messages = await graph.getMessages(
        req.app.locals.msalClient,
        req.session.userId,
        dateFns.formatISO(weekStart),
        dateFns.formatISO(weekEnd),
        user.timeZone
      );
      console.log(messages.value);

      // Assign the messages to the view parameters
      params.messages = messages.value;
      // res.status(200).json({
      //   message: "success",
      //   data: params,
      // });
    } catch (err) {
      console.log(err);
      req.flash("error_msg", {
        message: "Could not fetch events",
        debug: JSON.stringify(err, Object.getOwnPropertyNames(err)),
      });
    }

    res.render("mail", params);
  }
});

router.get("/details/:id", async function (req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect("/");
  } else {
    const params = {
      // active: { calendar: true },
      active: { mail: true },
    };

    // Get the user
    const user = req.app.locals.users[req.session.userId];
    // Convert user's Windows time zone ("Pacific Standard Time")
    // to IANA format ("America/Los_Angeles")
    const timeZoneId = iana.findIana(user.timeZone)[0];
    console.log(`Time zone: ${timeZoneId.valueOf()}`);

    // Calculate the start and end of the current week
    // Get midnight on the start of the current week in the user's timezone,
    // but in UTC. For example, for Pacific Standard Time, the time value would be
    // 07:00:00Z
    var weekStart = zonedTimeToUtc(
      dateFns.startOfWeek(new Date()),
      timeZoneId.valueOf()
    );
    var weekEnd = dateFns.addDays(weekStart, 7);
    console.log(`Start: ${dateFns.formatISO(weekStart)}`);

    try {
      // Get the events

      const messages = await graph.getMessagesDetails(
        req.app.locals.msalClient,
        req.session.userId,
        req.params.id,
        user.timeZone
      );
      console.log(messages);

      // Assign the messages to the view parameters
      params.messages = messages.value;
      // res.status(200).json({
      //   message: "success",
      //   data: params,
      // });
      res.send(messages.body.content);
    } catch (err) {
      console.log(err);
      req.flash("error_msg", {
        message: "Could not fetch events",
        debug: JSON.stringify(err, Object.getOwnPropertyNames(err)),
      });
    }

    // res.render("mail", params);
  }
});

module.exports = router;
