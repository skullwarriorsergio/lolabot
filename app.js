require("dotenv").config({ path: __dirname + "/.env" })
const { Telegraf, Markup } = require("telegraf")
const Excel = require('exceljs')
const fs = require("fs")
const {download} = require('./download')
const fileExists = require('./fsHelpers')
const { time } = require("console")
var bot = null
var options = [];
var stoppingBot = false
bot = new Telegraf(process.env.token)

//  download function
function downloadFiles(){
  try {
    download(process.env.excelfburl,process.env.excelfbfile)
  } catch (error) {
  }
  try {
    download(process.env.excelaccounturl,process.env.excelaccountfile)
  } catch (error) {
  }
  try {
    download(process.env.excellvrurl,process.env.excellvrfile)
  } catch (error) {
  }
  try {
    download(process.env.excelmfsurl,process.env.excelmfsfile)
  } catch (error) {
  }
}
//  Loop function
function downloadLoop(){
  setTimeout(function () {
    downloadFiles()
    if (!stoppingBot)
      downloadLoop();
    else console.log("Loop stopped")
  }, 400000) //400 000
}

//Execute download and start loop
downloadFiles()
downloadLoop()

bot.telegram.getMe().then((botInfo) => {
    bot.options.username = botInfo.username
  })
 
var workbook = new Excel.Workbook();

//-----------Code-----------
bot.command("start", (ctx) => {
  Welcome(ctx)
})
bot.command("lola", (ctx) => {
  Welcome(ctx)
})
bot.command("negativeslvr", (ctx) => {
  fileExists(process.env.excellvrfile).then((result) => {  
    if (!result)
    {
      ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
      return;    
    }
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    .then(() =>{
      bot.telegram.sendMessage(ctx.from.id,"Buscando <b>saldos negativos</b> en LVR. Espere por favor",{ parse_mode: 'HTML' }).then(() => {
        workbook.xlsx.readFile(process.env.excellvrfile).catch((err) => {})
        .then(function() { 
          //iterar por cada pagina
          workbook.eachSheet(function(worksheet, sheetId) {
            //iterar por cada file
            worksheet.eachRow(function(row, rowNumber) {
              if ((String(row.getCell(2).value).toLowerCase() === 'negative' || String(row.getCell(2).value).toLowerCase() === 'negativo') && row.getCell(5).value === null){
                bot.telegram.sendMessage(ctx.from.id,`• ${worksheet.name}    - <b>${row.getCell(4).value}</b>`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {
                }, 100))
              }
            })
          })
        })
      })
    })
  })
})
bot.command("negativesmfs", (ctx) => {
  fileExists(process.env.excelmfsfile).then((result) => {  
    if (!result)
    {
      ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
      return;    
    }
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    .then(() =>{
      bot.telegram.sendMessage(ctx.from.id,"Buscando <b>saldos negativos</b> en MFS. Espere por favor",{ parse_mode: 'HTML' }).then(() => {
        workbook.xlsx.readFile(process.env.excelmfsfile).catch((err) => {})
        .then(function() { 
          //iterar por cada pagina
          workbook.eachSheet(function(worksheet, sheetId) {
            //iterar por cada file
            worksheet.eachRow(function(row, rowNumber) {
              if ((String(row.getCell(2).value).toLowerCase() === 'negative' || String(row.getCell(2).value).toLowerCase() === 'negativo') && row.getCell(5).value === null){
                bot.telegram.sendMessage(ctx.from.id,`• ${worksheet.name}    - <b>${row.getCell(4).value}</b>`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {
                }, 100))
              }
            })
          })
        })
      })
    })
  })
})
//forzar actualizacion
bot.command("update", (ctx) => {
  ctx.reply("Actualizando excels desde Google Drive...")
  .then(() =>{
    downloadFiles()
  }).then(ctx.reply("Actualización completada."))
})
//Ingresos pendientes
bot.command("incomes", (ctx) => {
  fileExists(process.env.excelaccountfile).then((result) => {  
    if (!result)
    {
      ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
      return;    
    }
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    .then(() =>{
      bot.telegram.sendMessage(ctx.from.id,"Buscando <b>ingresos pendientes</b>. Espere por favor",{ parse_mode: 'HTML' }).then(() => {          
        workbook.xlsx.readFile(process.env.excelaccountfile)
        .catch((err) => {})
        .then(function() {          
            let worksheet = workbook.getWorksheet('Ingresos LVR')
            worksheet.eachRow(function(row, rowNumber) {
              if (row.getCell(3).value !== null && row.getCell(6).value === null){
                bot.telegram.sendMessage(ctx.from.id,`• LVR • Deudor <b>${row.getCell(5).value}</b> monto: <b>${row.getCell(4).value}</b>\n referencia: <b>${row.getCell(8).value}</b>`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {
                }, 100))
              }
            })
            let worksheet2 = workbook.getWorksheet('Ingresos MFS')
            worksheet2.eachRow(function(row, rowNumber) {
              if (row.getCell(3).value !== null && row.getCell(6).value === null){
                bot.telegram.sendMessage(ctx.from.id,`• MFS • Deudor <b>${row.getCell(5).value}</b> monto: <b>${row.getCell(4).value}</b>\n referencia: <b>${row.getCell(8).value}</b>`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {
                }, 100))
              }
            })
            let worksheet3 = workbook.getWorksheet('Ingresos HotShot')
            worksheet3.eachRow(function(row, rowNumber) {
              if (row.getCell(3).value !== null && row.getCell(6).value === null){
                bot.telegram.sendMessage(ctx.from.id,`• Deudor <b>${row.getCell(5).value}</b> monto: <b>${row.getCell(4).value}</b>\n referencia: <b>${row.getCell(8).value}</b>`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {
                }, 100))
              }
            })  
          })
      }).catch((err) =>{
        if (err.code === 403) {
          bot.telegram.sendMessage(
            -507850928,
            "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
          )
        }
      })
    })
  })
})
//Pagos pendientes
bot.command("payments", (ctx) => {
  fileExists(process.env.excelaccountfile).then((result) => {  
    if (!result)
    {
      ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
      return;    
    }
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    .then(() =>{
      bot.telegram.sendMessage(ctx.from.id,"Buscando <b>pagos pendientes</b>. Espere por favor",{ parse_mode: 'HTML' }).then(() => {          
        workbook.xlsx.readFile(process.env.excelaccountfile)
        .then(function() {           
          const worksheet = workbook.getWorksheet('Obligaciones de Pago')
          worksheet.eachRow(function(row, rowNumber) {
            if (row.getCell(3).value !== null && row.getCell(6).value === null){
              bot.telegram.sendMessage(ctx.from.id,`• Pagar a <b>${row.getCell(4).value}</b> monto: <b>${row.getCell(5).value}</b>\n  referencia: <b>${row.getCell(8).value}</b>`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {
              }, 100))
            }
          })
        })
      }).catch((err) =>{
        if (err.code === 403) {
          bot.telegram.sendMessage(
            -507850928,
            "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
          )
        }
      })
    })
  })
})

// Cargas atrasadas segun fecha de entrega en el rate
bot.command("delay", (ctx) => {
  fileExists(process.env.excelfbfile).then((result) => {  
    if (!result)
    {
      ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
      return;    
    }
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...") 
    bot.telegram.sendMessage(ctx.from.id,"Buscando cargas <b>atrasadas</b>. Espere por favor",{ parse_mode: 'HTML' }).then(() => {          
      workbook.xlsx.readFile(process.env.excelfbfile)
      .then(function() {           
          loadsDelayed= []
          loads = ""
          const worksheet = workbook.getWorksheet('Loads')
          statusCol = worksheet.getColumn('A')
          statusCol.eachCell(function(cell, rowNumber) {
              if (cell.text === 'ATRASADO'){
                  loadsDelayed.push(rowNumber)
              }
            });
          })
          .catch((err) => bot.telegram.sendMessage(ctx.from.id,"Un momento por favor, se están actualizando los datos"))
          .then(() =>
          bot.telegram.sendMessage(ctx.from.id,"Se encontraron <b>" + loadsDelayed.length + "</b> cargas <b>atrasadas</b>. Obteniendo detalles...",{ parse_mode: 'HTML' })).then(() =>{
          loadsDelayed?.forEach(element => {
              const worksheet = workbook.getWorksheet('Loads')
              rowData = worksheet.getRow(element)
              bot.telegram.sendMessage(ctx.from.id,`• Carga: <b>${rowData.getCell(6)}</b> Broker: <b>${rowData.getCell(7)}</b> Camión: <b>${rowData.getCell(8)}</b>\n(<b>${rowData.getCell(11)} ==> <b>${rowData.getCell(13)}</b></b>)\nDebió entregar el <b>${rowData.getCell(14).toString().slice(3, 15)}</b>`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {                
              }, 100))
        })
      })
    }).catch((err) =>{
      if (err.code === 403) {
        bot.telegram.sendMessage(
          -507850928,
          "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
        )
      }
    })
  })
})

//Cargas pendientes de pago por Factory / Broker
bot.command("pending", (ctx) => {
  fileExists(process.env.excelfbfile).then((result) => {  
    if (!result)
    {
      ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
      return;    
    }
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    .then(() => {
      bot.telegram.sendMessage(ctx.from.id,"Buscando cargas <b>pendientes de pago</b>. Espere por favor...",{ parse_mode: 'HTML' }).then(() => {      
        workbook.xlsx.readFile(process.env.excelfbfile)
        .catch((err) => bot.telegram.sendMessage(ctx.from.id,"Un momento por favor, se están actualizando los datos"))
        .then(function() {      
            var loads =""
            const worksheet = workbook.getWorksheet('Loads')
            worksheet.eachRow(function(row, rowNumber) {
            if (row.getCell(1).value?.result == 'BOL recibido' && row.getCell(4).value !== null && row.getCell(5).value === null)
                {
                  bot.telegram.sendMessage(ctx.from.id,`• Carga: <b>${row.getCell(6)}</b> Broker: <b>${row.getCell(7)}</b> Camión: <b>${row.getCell(8)}</b> Linehaul: <b>${row.getCell(10)}</b>\n(<b>${row.getCell(11)} ==> <b>${row.getCell(13)}</b></b>)`,{ parse_mode: 'HTML' })
                  .then(() => setTimeout(() => {
                  
                  }, 100))
                }
            });        
          })    
      }).catch((err) =>{
        if (err.code === 403) {
          bot.telegram.sendMessage(
            -507850928,
            "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
          )
        }
      })
    })
  })
})

// Cargas en HOLD en el factory
bot.command("hold", (ctx) => {  
  fileExists(process.env.excelfbfile).then((result) => {  
    if (!result)
    {
      ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
      return;
    }
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    .then(() => {
      bot.telegram.sendMessage(ctx.from.id,"Buscando cargas en <b>HOLD</b>. Espere por favor...",{ parse_mode: 'HTML' }).then(() => {      
        workbook.xlsx.readFile(process.env.excelfbfile)
        .catch((err) => bot.telegram.sendMessage(ctx.from.id,"Un momento por favor, se están actualizando los datos"))
        .then(function() {
            loadsDelayed= []
            const worksheet = workbook.getWorksheet('Issues')
            statusCol = worksheet.getColumn('C')
            statusCol.eachCell(function(cell, rowNumber) {
                if (cell.text === 'Hold'){
                    loadsDelayed.push(rowNumber)
                }
              });
              bot.telegram.sendMessage(ctx.from.id,"Existe(n) <b>" + loadsDelayed.length + "</b> carga(s) en <b>HOLD</b> en el factory. Obteniendo detalles...",{ parse_mode: 'HTML' })
              .then(() => {
                loadsDelayed?.forEach(element => {
                rowData = worksheet.getRow(element)
                bot.telegram.sendMessage(ctx.from.id,`• Carga: <b>${rowData.getCell(5)}</b> Broker: <b>${rowData.getCell(6)}</b> Camión: <b>${rowData.getCell(4)}</b> Motivo: <b>${rowData.getCell(7)}</b>`,{ parse_mode: 'HTML' })
                .then(() => setTimeout(() => {
                  
                }, 150))
              })
            })
          })
        }).catch((err) => {
          if (err.code === 403) {
            bot.telegram.sendMessage(
              -507850928,
              "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
            )
          }
        })
      })
    })
  })


function Welcome(ctx) {
    if (ctx.chat.id === ctx.from.id) {
      MainMenu(ctx);
    } else {
      var msg = `Hola <strong>${ctx.from.first_name}</strong>!\n`;
      msg += "Soy bot de ayuda de <strong>Lola VATB</strong>.\n";
      ctx.replyWithHTML(msg);
    }
}

function MainMenu(ctx) {
    let privateMSG = `Bienvenido <b>${ctx.from.first_name}</b>\nGracias por permitirme ayudarte. Que información necesitas?\nUtiliza los comandos para obtener lo que buscas.`;
    return ctx.replyWithHTML(privateMSG);
  }

bot.action("start", (ctx) => {
    var msg = `Hola <strong>${ctx.from.first_name}</strong>!\n`;
      msg += "Soy bot de ayuda de <strong>Lola VATB</strong>.\n";
      ctx.replyWithHTML(msg);
});

bot.on('callback_query', (ctx) => {
    ctx.answerCbQuery()
  })
  
bot.on('inline_query', (ctx) => {
    const result = []
    ctx.answerInlineQuery(result)
})
bot.catch((err) => {
  //console.log("bot error: ", err);
  if (
    err.code === 403 &&
    err.description.includes("bot was blocked by the user")
  ) {
    bot.telegram.sendMessage(
      -507850928,
      "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
    )
  }
  if (
    err.code === 403 &&
    err.description.includes("Forbidden: bot can't initiate conversation with a user")
  ) {
    bot.telegram.sendMessage(
      -507850928,
      "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
    )
  }
})
  
bot.launch()
  
// Enable graceful stop
process.once('SIGINT', () => {
  stoppingBot=true
  bot.stop('SIGINT')  
})
process.once('SIGTERM', () => {
  stoppingBot=true
  bot.stop('SIGTERM')
})