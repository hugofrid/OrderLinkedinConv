
const readXlsxFile = require('read-excel-file/node');
const fs = require('fs');
const xlsx = require("node-xlsx");





function isSameCouple(coupl1, coupl2) {

  
    
    //console.log(coupl1[0], coupl2[0], coupl1[1] , coupl2[1])

    if ((coupl1[0] === coupl2[0] && coupl1[1]===coupl2[1]) ||
        (coupl1[0] === coupl2[1] && coupl1[1]===coupl2[0]) ||
        (coupl1[1] === coupl2[0] && coupl1[0]===coupl2[1])) {
       
        return true
    }if (coupl1[1].split(',').length > 1 && coupl2[1].split(',').length > 1) {
        if (coupl1[0] === coupl2[1].split(',')[0] && coupl1[1] === coupl2[1].replace(coupl2[1].split(',')[0],coupl2[0])) {
         
            return true;
           

        }
        if (coupl2[0] === coupl1[1].split(',')[0] && coupl2[1] === coupl1[1].replace(coupl1[1].split(',')[0],coupl1[0])) {
           
            return true;
        }
    }
    else return false

}

function compare(a, b) {
    if (new Date(a[2]).getTime() < new Date(b[2]).getTime()) {
        return -1;
    }
    if (new Date(a[2]).getTime() > new Date(b[2]).getTime()) {
        return 1;
    }
    return 0;
}




// FROM | TO | DATA | SUBJECT | CONTENT | DIRECTION | FOLDER 
readXlsxFile('./files/messages-test.xlsx').then(async (rows) => {
    // `rows` is an array of rows
    // each row being an array of cells.
    let filteredMessages = [];
    let noAnswerFrom = [];
    try {
        console.log('start processing data');

        const messages = rows.slice(1);

        await messages.map(message => {
            let couple = [message[0], message[1]];


            // console.log('boucle map');

            let conv = [...messages.filter(msg => isSameCouple(couple, [msg[0], msg[1]]) && (filteredMessages.findIndex(m => m === msg) === -1))];

            // console.log('message by same couple', conv.length, conv)
            conv.sort(compare);


            if (conv.findIndex(msg => msg[5] === 'OUTGOING') === -1) {
                filteredMessages.push(...conv, ['YOU DIDN\'T ANSWER THIS CONVERSATION'], [''])
            }
            if (conv.findIndex(msg => msg[5] === 'INCOMING') === -1) {
                filteredMessages.push(...conv, ['YOUR CORRESPONDANT DIDN\'T ANSWER THIS CONVERSATION'], [''])
            }
            else {
                filteredMessages.push(...conv,[''])
            
            }
            // await console.log(conv.length,filteredMessages.length)
    

        })



        const buffer = xlsx.build([{ name: "demo_sheet", data: filteredMessages }])
        fs.writeFile('./outputFile/result.xlsx', buffer, (err) => {
            if (err) throw err
            console.log('Done...');
        })

            await console.log('end of process');
        }
     catch (err) {
            console.log(err)
        }
    })





