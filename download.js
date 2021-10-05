const { http, https } = require('follow-redirects')
const fs = require("fs")

module.exports =  function download(url, filePath) {
  const file = fs.createWriteStream(filePath);
  const request = https.request(url, response => {
    response.pipe(file);
  });
  request.end();
}