import * as OfficeHelpers from '@microsoft/office-js-helpers';

$(document).ready(() => {
    $('#run').click(run);
});

// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

async function run() {
    try {
        await Excel.run(async context => {
            var range = context.workbook.worksheets.getActiveWorksheet().getRange("A:AE").getUsedRange();

            var afex = context.workbook.worksheets.add("Afex2");
            range.load("values");
            await context.sync();
            var rangeValues = range.values;

            var BpCodeIndex = rangeValues[0].indexOf("BP Code"); 
            var CurrencyIndex = rangeValues[0].indexOf("Payment Method"); 
            var ValueDateIndex = rangeValues[0].indexOf("Due Date"); 
            var AmountIndex = rangeValues[0].indexOf("Pmt Amount"); 


            rangeValues = rangeValues.filter(x => x[0] !== "");
            var cleanRange = [];
            rangeValues.forEach(row => {

                row = [
                    row[CurrencyIndex].substring(5,8),
                    row[AmountIndex+1] === "" ? row[AmountIndex] : row[AmountIndex+1].substring(4),
                    row[ValueDateIndex],
                    row[BpCodeIndex]
                ];

                cleanRange.push(row);
            });

            cleanRange[0] = ["Currency", "Amount", "Value Date", "Vendor ID"];       

            var i = 1;
            do {
                console.log("cell " + cleanRange[i][3]);
                console.log("cell " + JSON.stringify(cleanRange));
                var cell = cleanRange[i][3];
                var nextCell = cleanRange[i+1][3];

                if (cell === nextCell && cell !== "") {
                    cleanRange[i][1] = cleanRange[i][1] + cleanRange[i + 1][1];
                    cleanRange[i][2] = Math.min(cleanRange[i][2], cleanRange[i + 1][2]);
                   
                    cleanRange.splice(i+1, 1);
                } else {
                    i++;
                }
            } while (i <cleanRange.length - 1);

            var targetRange = afex.getRange("A1:D"+cleanRange.length.toString());
            targetRange.set({values: cleanRange});

            targetRange.getColumn(2).numberFormat = "mm/dd/yyy";
            targetRange.getColumn(1).numberFormat = "0,0.00";

            targetRange.load();
            await context.sync();

            var ch = "";
            targetRange.text.forEach(row => {
                if (row[1] !== "Amount"){
                    row[1] = '"' + row[1] + '"';
                }
                
                ch += row.join("\t"); 
                ch += '\n';
            });

            var FileSaver = require('file-saver');
            var blob = new Blob([ch], { type: "text/plain;charset=utf-8" });
            FileSaver.saveAs(blob, "afex.txt");
       })
    } catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    };
}