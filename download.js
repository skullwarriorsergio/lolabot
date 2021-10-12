require("dotenv").config({ path: __dirname + "/.env" })
const { http, https } = require('follow-redirects')
const fs = require("fs")

const download = (url, filePath, callback) => {
  try {
    const file = fs.createWriteStream(filePath,{flags: 'w'});
    const request = https.request(url, response => {
      response.pipe(file);
    });
    request.end();
    file.on('finish', function(){
      console.log('finished downloading '+ filePath);    
      file.close()
      if (callback)
          callback(filePath)
    });
  } catch (error) {
    console.log(error)
  }
}

module.exports = { download }