require("dotenv").config({ path: __dirname + "/.env" })
const { Telegraf, Markup } = require("telegraf")
const Excel = require('exceljs')
const fs = require("fs")
const download = require('./download')
const { time } = require("console")
var bot = null
var options = [];
var stoppingBot = false
bot = new Telegraf(process.env["token"])

//Download Helpers
function timeout() {
  setTimeout(function () {
    download("https://drive.google.com/uc?id=1-bDcHJFL9Otj_HmF--5KZsrcRasb72PL&export=download","file.xlsx")
    console.log("adsad")
    if (!stoppingBot)
      timeout();
  }, 500000)
}

//Execute the downloaders
timeout()
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
// Cargas atrasadas segun fecha de entrega en el rate
bot.command("delay", (ctx) => {  
  bot.telegram.sendMessage(ctx.from.id,"Buscando cargas <b>atrasadas</b>. Espere por favor",{ parse_mode: 'HTML' }).then(() => { 
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    workbook.xlsx.readFile("file.xlsx")
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
        }).then(() =>
        bot.telegram.sendMessage(ctx.from.id,"Se encontraron <b>" + loadsDelayed.length + "</b> cargas <b>atrasadas</b>. Obteniendo detalles...",{ parse_mode: 'HTML' })).then(() =>{
        loadsDelayed?.forEach(element => {
            const worksheet = workbook.getWorksheet('Loads')
            rowData = worksheet.getRow(element)
            bot.telegram.sendMessage(ctx.from.id,"Carga: <b>"+ rowData.getCell(6) + "</b> Broker: <b>" + rowData.getCell(7) + "</b> Camión: <b>" + rowData.getCell(8)+ "</b> Debió entregar el <b>" + rowData.getCell(14).toString().slice(3, 15) + "</b>\n",{ parse_mode: 'HTML' })
      })
    })
  }).catch((err) =>{
    bot.telegram.sendMessage(
      -507850928,
      "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
    );
  })
})

//Cargas pendientes de pago por Factory / Broker
bot.command("pending", (ctx) => {
  bot.telegram.sendMessage(ctx.from.id,"Buscando cargas <b>pendientes de pago</b>. Espere por favor...",{ parse_mode: 'HTML' }).then(() => {
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    workbook.xlsx.readFile("file.xlsx")
    .then(function() {      
        var loads =""
        const worksheet = workbook.getWorksheet('Loads')
        worksheet.eachRow(function(row, rowNumber) {
        if (row.getCell(1).value?.result == 'BOL recibido' && row.getCell(4).value !== null && row.getCell(5).value === null)
            {
              bot.telegram.sendMessage(ctx.from.id,"Carga: <b>"+ row.getCell(6) + "</b> Broker: <b>" + row.getCell(7) + "</b> Camión: <b>" + row.getCell(8)+ "</b> Linehaul: <b>" + row.getCell(10) + "</b>\n",{ parse_mode: 'HTML' })
              .then(() => setTimeout(() => {
              
              }, 100))
            }
        });        
      })    
  }).catch((err) =>{
    bot.telegram.sendMessage(
      -507850928,
      "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
    );
  })
})

// Cargas en HOLD en el factory
bot.command("hold", (ctx) => {    
  bot.telegram.sendMessage(ctx.from.id,"Buscando cargas en <b>HOLD</b>. Espere por favor...",{ parse_mode: 'HTML' }).then(() => {
    ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
    workbook.xlsx.readFile("file.xlsx")
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
            bot.telegram.sendMessage(ctx.from.id,"Carga: <b>"+ rowData.getCell(5) + "</b> Broker: <b>" + rowData.getCell(6) + "</b> Camión: <b>" + rowData.getCell(4) + "</b> Motivo: <b>" + rowData.getCell(7) +"</b>",{ parse_mode: 'HTML' })
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
    let privateMSG = `Bienvenido ${ctx.from.first_name}\nGracias por permitirme ayudarte.\nQue deseas hacer?.`;
    return MainMenuButtons(ctx, privateMSG);
  }
function MainMenuButtons(ctx, menuMSG) {
    return bot.telegram.sendMessage(ctx.from.id, menuMSG, {
      reply_markup: {
        inline_keyboard: [
          [
            {
              text: "\u{26D1} Iniciar asistente",
              callback_data: "start",
            },
          ],
        ],
      },
    });
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