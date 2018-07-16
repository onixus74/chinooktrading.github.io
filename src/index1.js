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
            var range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:M50").getUsedRange();

            var afex = context.workbook.worksheets.add("Afex");
            range.load("values");
            await context.sync();
            var rangeValues = range.values;

            var BpCodeIndex = rangeValues[0].indexOf("Bp Code"); 
            var CurrencyIndex = rangeValues[0].indexOf("Currency"); 
            var ValueDateIndex = rangeValues[0].indexOf("Due Date"); 
            var AmountIndex = rangeValues[0].indexOf("Credit Amount"); 


            rangeValues = rangeValues.filter(x => x[0] !== "");
            var cleanRange = [];
            rangeValues.forEach(row => {
                row = [row[CurrencyIndex],row[AmountIndex],row[ValueDateIndex],row[BpCodeIndex]];
                cleanRange.push(row);
            });

            cleanRange[0] = ["Currency", "Amount", "Value Date", "Vendor ID"];       

            var targetRange = afex.getRange("A1:D"+cleanRange.length.toString());
            targetRange.set({values: cleanRange});

            targetRange.getColumn(2).numberFormat = "mm/dd/yyy";
            targetRange.getColumn(1).numberFormat = "0,0.00";

            targetRange.load();
            await context.sync();

            var i = 1;
            do {
                var cell = targetRange.getCell(i, 3);
                var nextCell = targetRange.getCell(i + 1, 3);
                cell.load();
                nextCell.load();
                await context.sync();

                if (cell.values[0][0] === nextCell.values[0][0] && cell.values[0][0] !== "") {
                    var firstCell = targetRange.getCell(i, 1);
                    var secondCell = targetRange.getCell(i + 1, 1);
                    firstCell.load("values");
                    secondCell.load("values");
                    await context.sync();

                    firstCell.set({ values: firstCell.values[0][0] + secondCell.values[0][0] });
                    secondCell.format.fill.color = "green";

                    firstCell = targetRange.getCell(i, 2);
                    secondCell = targetRange.getCell(i + 1, 2);

                    firstCell.load("values");
                    secondCell.load("values");
                    await context.sync();

                    firstCell.set({ values: Math.min(firstCell.values[0][0], secondCell.values[0][0]) });

                    nextCell.getEntireRow().delete();
                } else {
                    i++;
                    targetRange.getUsedRange().load();
                    await context.sync();
                }
            } while (i <= targetRange.rowCount - 1);

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