require("dotenv").config({ path: __dirname + "/.env" })
const { http, https } = require('follow-redirects')
const fs = require("fs")

function download(url, filePath)  {
  const file = fs.createWriteStream(filePath);
  const request = https.request(url, response => {
    response.pipe(file);
  });
  request.end();
}

function timeoutExcelFamilyBussiness() {
  setTimeout(function () {
    download(process.env.excelfburl,process.env.excelfbfile)
    if (!stoppingBot)
      timeout();
  }, 500000)
}

module.exports = { download, timeoutExcelFamilyBussiness }