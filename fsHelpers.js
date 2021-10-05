const fs = require("fs")

const fileExists = (file) => {
    return new Promise((resolve) => {
        fs.access(file, fs.constants.F_OK, (err) => {
            err ? resolve(false) : resolve(true)
        });
    })
}

module.exports = function checkFileExists(path){
    return fileExists(path)
}

