
const fileExists = require('./fsHelpers')
let dollarUSLocale = Intl.NumberFormat('en-US', {
    minimumFractionDigits: 2,      
    maximumFractionDigits: 2,
  });
const Excel = require('exceljs')
const workbook = new Excel.Workbook();

//Get negatives
const getNegatives = async (bot, ctx, file, title) => {
    try {
        const exist = await fileExists(file)
        if (!exist)
        {
            await ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
            return;
        }
        await ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
        await bot.telegram.sendMessage(ctx.from.id,`Buscando <b>saldos negativos</b> en ${title}. Espere por favor\nTenga en cuenta que si existen payrolls en revisión, los negativos que puedan surgir no se mostrarán hasta que se aprueben dichos payrolls`,{ parse_mode: 'HTML' })
        await workbook.xlsx.readFile(file).catch((err) => {})
        //iterar por cada pagina
        await workbook.eachSheet(async function(worksheet, sheetId) {
            //iterar por cada file
            await worksheet.eachRow(async function(row, rowNumber) {
              if ((String(row.getCell(2).value).toLowerCase() === 'negative' || String(row.getCell(2).value).toLowerCase() === 'negativo') && row.getCell(5).value === null){
                await bot.telegram.sendMessage(ctx.from.id,`• ${worksheet.name}    - <b>$${dollarUSLocale.format(row.getCell(4).value)}</b>`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {
                }, 100))
              }
            })
          })
        await ctx.replyWithHTML("______Consulta completada______")

    } catch (error) {
        console.log(error)
    }
}

const getIncomes = async (bot, ctx) => {
    try {
        const result = await fileExists(process.env.excelaccountfile)
        if (!result)
        {
            await ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
            return;    
        }        
        await ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
        await bot.telegram.sendMessage(ctx.from.id,"Buscando <b>ingresos pendientes</b>. Espere por favor",{ parse_mode: 'HTML' })
        await workbook.xlsx.readFile(process.env.excelaccountfile)
        let worksheet = workbook.getWorksheet('Ingresos LVR')
        worksheet.eachRow(async (row, rowNumber) => {
            if (row.getCell(3).value !== null && row.getCell(6).value === null){
                await bot.telegram.sendMessage(ctx.from.id,`• LVR • Deudor <b>${row.getCell(5).value}</b> monto: <b>$${dollarUSLocale.format(row.getCell(4).value)}</b>\n referencia: <b>${row.getCell(8).value}</b>`,{ parse_mode: 'HTML' })
            }
        })
        let worksheet2 = workbook.getWorksheet('Ingresos MFS')
        worksheet2.eachRow(async (row, rowNumber) => {
            if (row.getCell(3).value !== null && row.getCell(6).value === null){
                await bot.telegram.sendMessage(ctx.from.id,`• MFS • Deudor <b>${row.getCell(5).value}</b> monto: <b>$${dollarUSLocale.format(row.getCell(4).value)}</b>\n referencia: <b>${row.getCell(8).value}</b>`,{ parse_mode: 'HTML' })
            }
        })
        let worksheet3 = workbook.getWorksheet('Ingresos HotShot')
        worksheet3.eachRow(async (row, rowNumber) => {
          if (row.getCell(3).value !== null && row.getCell(6).value === null){
                await bot.telegram.sendMessage(ctx.from.id,`• Deudor <b>${row.getCell(5).value}</b> monto: <b>$${dollarUSLocale.format(row.getCell(4).value)}</b>\n referencia: <b>${row.getCell(8).value}</b>`,{ parse_mode: 'HTML' })
          }
        })

    } catch (error) {
        if (error.code === 403) {
            bot.telegram.sendMessage(
              -507850928,
              "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
            )
          }
    }
}

const getPayments = async (bot, ctx) =>{
    try {
        const result = await fileExists(process.env.excelaccountfile)
        if (!result)
        {
          await ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
          return;    
        }
        await ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
        await bot.telegram.sendMessage(ctx.from.id,"Buscando <b>pagos pendientes</b>. Espere por favor",{ parse_mode: 'HTML' })
        await workbook.xlsx.readFile(process.env.excelaccountfile)
        const worksheet = workbook.getWorksheet('Obligaciones de Pago')
        worksheet.eachRow(async (row, rowNumber) => {
            if (row.getCell(3).value !== null && row.getCell(6).value === null){
              await bot.telegram.sendMessage(ctx.from.id,`• Pagar a <b>${row.getCell(4).value}</b> monto: <b>$${dollarUSLocale.format(row.getCell(5).value)}</b>\n  referencia: <b>${row.getCell(8).value}</b>`,{ parse_mode: 'HTML' })
            }
          })
    } catch (error) {
        if (error.code === 403) {
            bot.telegram.sendMessage(
              -507850928,
              "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
            )
          }
    }
}

const getDelay = async (bot,ctx) =>{
    try {
        const result = await fileExists(process.env.excelfbfile)
        if (!result)
        {
            await ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
            return;    
        }
        await ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...") 
        await bot.telegram.sendMessage(ctx.from.id,"Buscando cargas <b>atrasadas</b>. Espere por favor",{ parse_mode: 'HTML' })
        await workbook.xlsx.readFile(process.env.excelfbfile)
        let loadsDelayed = []
        let loads = ""
        const worksheet = workbook.getWorksheet('Loads')
        statusCol = worksheet.getColumn('A')
        statusCol.eachCell((cell, rowNumber) => {
            if (cell.text === 'ATRASADO'){
                loadsDelayed.push(rowNumber)
            }
        })
        await bot.telegram.sendMessage(ctx.from.id,"Se encontraron <b>" + loadsDelayed.length + "</b> cargas <b>atrasadas</b>. Obteniendo detalles...",{ parse_mode: 'HTML' })
        loadsDelayed?.forEach(async (element) => {
            const worksheet = workbook.getWorksheet('Loads')
            rowData = worksheet.getRow(element)
            await bot.telegram.sendMessage(ctx.from.id,`• Carga: <b>${rowData.getCell(6)}</b> Broker: <b>${rowData.getCell(7)}</b> Camión: <b>${rowData.getCell(8)}</b>\n(<b>${rowData.getCell(11)} ==> <b>${rowData.getCell(13)}</b></b>)\nDebió entregar el <b>${rowData.getCell(14).toString().slice(3, 15)}</b>`,{ parse_mode: 'HTML' })
        })             
    } catch (error) {
        if (error.code === 403) {
            bot.telegram.sendMessage(
              -507850928,
              "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
            )
          }
    }
}

const getPending = async (bot,ctx) => {
    let loads = 0
    try {
        const result = await fileExists(process.env.excelfbfile)
        if (!result)
        {
            await ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
            return;    
        }
        await ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
        await bot.telegram.sendMessage(ctx.from.id,"Buscando cargas <b>pendientes de pago</b>. Espere por favor...",{ parse_mode: 'HTML' })        
        await workbook.xlsx.readFile(process.env.excelfbfile)
        const worksheet = workbook.getWorksheet('Loads')
        worksheet.eachRow(async (row, rowNumber) => {
            loads++
            if (row.getCell(1).value?.result == 'BOL recibido' && row.getCell(4).value !== null && row.getCell(5).value === null)
                {
                    await bot.telegram.sendMessage(ctx.from.id,`• Carga: <b>${row.getCell(6)}</b> Broker: <b>${row.getCell(7)}</b> Camión: <b>${row.getCell(8)}</b> Monto: <b>$${dollarUSLocale.format(row.getCell(10))}</b>\n(<b>${row.getCell(11)} ==> <b>${row.getCell(13)}</b></b>)`,{ parse_mode: 'HTML' }).then(() => setTimeout(() => {                  
                    }, 150))
                }
        })
    } catch (error) {
        if (error.code === 403) {
            bot.telegram.sendMessage(
              -507850928,
              "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
            )
          }
    }
    finally
    {
        if (loads === 0)
            ctx.replyWithHTML(`No existen cargas pendientes`)
    }
}

const getHold = async (bot, ctx) =>{
    try {
        const result = await fileExists(process.env.excelfbfile)
        if (!result)
        {
          await ctx.replyWithHTML("Upss, el excel necesario no se encuentra, espere unos minutos a que se descargue o contacte con el desarrollador del bot.")
          return;
        }
        await ctx.replyWithHTML(ctx.chat.id != ctx.from.id ? `Hola <strong>${ctx.from.first_name}</strong>!\n` + "Le he enviado la respuesta a su consulta en un mensaje privado.\nNos vemos allí." : "Entendido, ejecutando comando...")
        await bot.telegram.sendMessage(ctx.from.id,"Buscando cargas en <b>HOLD</b>. Espere por favor...",{ parse_mode: 'HTML' })
        await workbook.xlsx.readFile(process.env.excelfbfile)
        let loadsDelayed = []
        const worksheet = workbook.getWorksheet('Issues')
        let statusCol = worksheet.getColumn('C')
        statusCol.eachCell(async (cell, rowNumber) =>{
            if (cell.text === 'Hold'){
                loadsDelayed.push(rowNumber)
            }
        })
        await bot.telegram.sendMessage(ctx.from.id,"Existe(n) <b>" + loadsDelayed.length + "</b> carga(s) en <b>HOLD</b> en el factory. Obteniendo detalles...",{ parse_mode: 'HTML' })
        loadsDelayed?.forEach(async (element) => {
            rowData = worksheet.getRow(element)
            await bot.telegram.sendMessage(ctx.from.id,`• Carga: <b>${rowData.getCell(5)}</b> Broker: <b>${rowData.getCell(6)}</b> Camión: <b>${rowData.getCell(4)}</b> Motivo: <b>${rowData.getCell(7)}</b>`,{ parse_mode: 'HTML' })
            .then(() => setTimeout(() => {
            }, 150))
        })

    } catch (error) {
        if (error.code === 403) {
            bot.telegram.sendMessage(
              -507850928,
              "Oh oh! No me ha sido posible enviarte un mensaje privado.\n Pudieras iniciar una conversación directa conmigo: @lolavatb_bot y ejecutar el comando /start"
            )        
        }
    }
}

module.exports = { getNegatives, getIncomes, getPayments, getDelay, getPending, getHold }