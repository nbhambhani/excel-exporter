
const express = require('express');
const app = express();
const generateExcel = require('./generateExcel');

app.get('/', function (req, res) {
    // Diplay a button and onClick, reroute to download
    res.send('Route to /download');
});

app.get('/download', function (req, res) {
  	
    generateExcel.createExcel();
  	const file =  '../data.xlsx';
  	const filename = 'volunteers.xlsx'
  	res.download(file, filename);
});

app.listen(3000, function () {
  	console.log('Example app listening on port 3000!');
});
