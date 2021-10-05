require("dotenv").config({ path: __dirname + "/.env" })
const { http, https } = require('follow-redirects')
const fs = require("fs")

function download(url, filePath)  {
  const file = fs.createWriteStream(filePath,{flags: 'w'});
  const request = https.request(url, response => {
    response.pipe(file);
  });
  request.end();
  file.on('finish', function(){
    console.log('finished downloading '+ filePath);
    file.close()
  });
}

module.exports = { download }