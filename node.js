var express = require('express');
var http = require('http');
var bodyParser = require('body-parser');
var multiparty = require('multiparty');
var xlsx = require('xlsx');
var fs = require('fs');
var mime = require('mime');

var app = express();

app.use(express.static(__dirname + '/public/html'));
app.use(express.static(__dirname + '/public/js'));

app.use(express.static('public'));

app.use(express.urlencoded());
app.use(express.json());

app.post('/exceltojson', function(req, res, _next) {
  var resData = {};
  var fileName = {};

  var form = new multiparty.Form({
      autoFiles: true
  });

  form.on('file', function(_name, file) {

    var splitdot = file.originalFilename.split(".");
    fileName = splitdot[0] + ".json";

      var workbook = xlsx.readFile(file.path);
      var sheetnames = Object.keys(workbook.Sheets);

      var i = sheetnames.length;

      while (i--) {
          var sheetname = sheetnames[i];
          resData[sheetname] = xlsx.utils.sheet_to_json(workbook.Sheets[sheetname]);
      }
  });

  form.on('close', function() {

  console.log(fileName);
  
  mimetype = mime.lookup(fileName);

  res.setHeader('Content-disposition', 'attachment; filename=' + fileName);
  res.setHeader('Content-type', mimetype);

  res.send(resData);

  });

  form.parse(req);
});

app.get('/index', function(_req, res) {

  res.writeHead(200,{'Content-Type':'text/html'});
  fs.readFile(__dirname + '/public/html/index.html', function(err, data) { 
    if (err) return console.error(err);
    
    res.end(data, 'utf-8');
  });
});

app.get('/exceltojson', function(_req, res) {

  res.writeHead(200,{'Content-Type':'text/html'});
  fs.readFile(__dirname + '/public/html/ExcelToJson.html', function(err, data) { 
    if (err) return console.error(err);
    
    res.end(data, 'utf-8');
  });
});

var host = 'localhost';
var port = 7777;

http.createServer(app).listen(port, host, function() {
  console.log('HTTP server listening on port ' + port);
});

