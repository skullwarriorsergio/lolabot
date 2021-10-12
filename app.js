require("dotenv").config({ path: __dirname + "/.env" })
const { Telegraf, Markup } = require("telegraf")
const Excel = require('exceljs')
const fs = require("fs")
const {download} = require('./download')
const { getIncomes, getNegatives, getPayments, getDelay, getPending, getHold } = require('./operations')
const fileExists = require('./fsHelpers')
var bot = null
var stoppingBot = false
bot = new Telegraf(process.env.token)
let dollarUSLocale = Intl.NumberFormat('en-US', {
  minimumFractionDigits: 2,      
  maximumFractionDigits: 2,
});
const workbook = new Excel.Workbook();

//  download function
const downloadFiles = (callback) => {
  try {
    download(process.env.excelfburl, process.env.excelfbfile, callback)
  } catch (error) {
  }
  try {
    download(process.env.excelaccounturl, process.env.excelaccountfile, callback)
  } catch (error) {
  }
  try {
    download(process.env.excellvrurl, process.env.excellvrfile, callback)
  } catch (error) {
  }
  try {
    download(process.env.excelmfsurl, process.env.excelmfsfile, callback)
  } catch (error) {
  }
}

//  Loop function
const downloadLoop = () => {
  setTimeout(() => {
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

//-----------Code-----------
bot.command("start", (ctx) => {
  Welcome(ctx)
})
bot.command("lola", (ctx) => {
  Welcome(ctx)
})
bot.command("negativeslvr", (ctx) => {
  getNegatives(bot, ctx, process.env.excellvrfile, "LVR")
})
bot.command("negativesmfs", (ctx) => {
  getNegatives(bot, ctx, process.env.excelmfsfile, "MFS") 
})
//forzar actualizacion
bot.command("update", async (ctx) => {
  await ctx.reply(`Se estan actualizando ${4} archivos desde google drive. Espere unos momentos por favor`)
  downloadFiles((file) => ctx.reply(`Actualización de ${file} completada.`)) 
})

//Ingresos pendientes
bot.command("incomes", (ctx) => {
  getIncomes(bot, ctx)
})

//Pagos pendientes
bot.command("payments", (ctx) => {
  getPayments(bot, ctx)
})

// Cargas atrasadas segun fecha de entrega en el rate
bot.command("delay", (ctx) => {
  getDelay(bot, ctx)
})

//Cargas pendientes de pago por Factory / Broker
bot.command("pending", (ctx) => {
  getPending(bot, ctx)
})

// Cargas en HOLD en el factory
bot.command("hold", (ctx) => {  
  getHold(bot, ctx)
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