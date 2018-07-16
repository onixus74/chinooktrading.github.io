import * as OfficeHelpers from '@microsoft/office-js-helpers';

$(document).ready(() => {
    $('#run').click(run);
});

// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

// var targetRange = context.workbook.worksheets.getItem("ApexTransactions").getRange("A1:M50").getUsedRange();
// Office.context.add(targetRange);

async function run() {
    try {
        await Excel.run(async context => {
            const columnToSave = ["Bp Code", "Due Date", "Credit Amount", "Currency"];

            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

            var apexTransactionSheet = context.workbook.worksheets.add("ApexTransactions");

            var range = currentWorksheet.getRange("A1:M50").getUsedRange();
            range.load("values");
            await context.sync();

            // console.log(`The range address was ${range.address}.`);

            var targetRange = apexTransactionSheet.getRange("A1:M50");

            targetRange.set({
                values: range.values
            })

            targetRange.load();
            await context.sync();

            for (let i = targetRange.rowCount - 1; i > 0; i--) {
                const row = targetRange.getRow(i);

                const cell = row.getCell(0, 0);

                cell.load("values");
                await context.sync();

                if (cell.values[0][0] === "") {
                    row.delete();
                }
            }

            for (let i = targetRange.columnCount - 1; i > 0; i--) {
                const col = targetRange.getColumn(i);
                const cell = col.getCell(0, 0)

                cell.load("values");
                await context.sync();

                if (columnToSave.indexOf(cell.values[0][0]) === -1) {
                    col.delete("Left");
                }
            }
        })
    } catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    };

    await Excel.run(async context => {
        try {

            var targetRange = context.workbook.worksheets.getItem("ApexTransactions").getRange("A1:M50").getUsedRange();
            targetRange.load();
            await context.sync();

            function switchColumn(a, b) {
                var col = targetRange.getColumn(a);
                var targetCol = targetRange.getColumn(b);
                col.load("values");
                targetCol.load("values");
                return context.sync().then(function () {
                    var t = col.values;
                    col.set({ values: targetCol.values });
                    targetCol.set({ values: t });
                });
            }

            switchColumn(0, 3).then(function () {
                switchColumn(1, 2);
            });

            var dateColumn = targetRange.getColumn(2);
            dateColumn.numberFormat = "mm/dd/yyy";
            await context.sync();

            var amountColumn = targetRange.getColumn(1);
            amountColumn.numberFormat = "0,0.00";
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

            var header = targetRange.getRow(0);
            header.load("values");
            await context.sync();

            header.set({
                values: [
                    ["Currency", "Amount", "Value Date", "Vendor ID"]
                ]
            });

            header.load("values");
            await context.sync();

            targetRange.load();
            await context.sync();

            console.log("Debug info: " + JSON.stringify(targetRange.values));

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
            FileSaver.saveAs(blob, "apex.txt");

        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    });
}