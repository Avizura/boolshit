'use strict';
/*
  Module dependencies.
 */
var express = require('express'),
  app = express(),
  server,
  server_port = '5000',
  // path = require('path'),
  appPath = process.cwd(),
  bodyParser = require('body-parser'),
  allowCrossDomain = function(req, res, next) {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE');
    res.header('Access-Control-Allow-Headers', 'Content-Type');
    next();
  };
var json2xls = require('json2xls');
var jsonArr = [{
  foo: 'bar',
  qux: 'moo',
  poo: 123,
  stux: new Date()
}, {
  foo: 'bar',
  qux: 'moo',
  poo: 345,
  stux: new Date()
}];

app.use(json2xls.middleware);

// app.get('/', function(res) {
//   res.xls('data.xlsx', jsonArr);
// });


app.use(allowCrossDomain);
app.use(bodyParser.urlencoded({
  extended: true
}));
app.use(bodyParser.json());

//Error handling
app.use(function(err, req, res, next) {
  console.error(err.stack);
});

// app.post('/feedback', function(req, res) {
//   console.log(req.body);
//   req.models.feedback.create({
//       login: req.session.login,
//       feedback_type: req.body.feedbackType,
//       text: req.body.text
//     },
//     function(err, items) {
//       console.log(err);
//       res.end();
//     });
// });

app.post('/mytest', function(req, res) {
  console.log(req.body);
  console.log("AVIZURA AAA");
  // res.end("Everything is good!");
  res.xls('data.xlsx', jsonArr);
});


server = app.listen(server_port, function() {
  console.log("Listening on port " + server_port);
});
