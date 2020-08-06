function onOpen(e) {
    let menu = SpreadsheetApp.getUi().createMenu('Custom Functions');
    menu.addItem('Update and Transfer Data', 'updateTransferData')
        .addItem('Import Data', 'importCsv')
        .addItem('Transfer Data', 'transferQueryData')
        .addItem('Run Queries', 'fillQueries')
        .addToUi();
}

function updateTransferData() {
    importCsv();
    fillQueries();
    transferQueryData();
}

function importCsv () {
    const dataImport = new ImportData();
    dataImport.importCSVData();
}

function columnToLetter(columnNum) {
    let temp, letter = '';
    while (columnNum > 0) {
        temp = (columnNum - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        columnNum = (columnNum - temp - 1) / 26;
    }
    return letter;
}


class ImportData {
    constructor() {
        this.sheet = SpreadsheetApp.openById("1jPFeGJgeU0nyHZx9DXTL0wagIQM3Cbl5AyzD37ky3ck")
            .getSheetByName('API Overview - For Data Studio');
        this.sheetData = this.sheet.getDataRange().getValues();
        this.dataUrl = "https://www.collegeoverview.com/common/billing/csv_posting/";
        this.csvData = this.getCsvData();
    }
    importCSVData() {
        const numCols = this.csvData[0].length;
        const lastColLetter = columnToLetter(numCols);
        const importRange = 'A1:' + lastColLetter + this.csvData.length;

        this.sheet.getDataRange().clear();
        this.sheet.getRange(importRange).setValues(this.csvData);
        this.formatOverviewCells();
    }
    getCsvData() {
        // TODO can update this to login to posting page and put it behind a login
        const urlContent = UrlFetchApp.fetch(this.dataUrl).getContentText();
        return Utilities.parseCsv(urlContent);
    }
    formatOverviewCells() {
        for (let i = 0; i < this.sheetData[0].length; i++) {
            const colLetter = columnToLetter(i+1);

            let format;

            const header = this.sheetData[0][i];
            const valType = typeof this.sheetData[1][i];
            if (valType === 'number' && header === 'Sold') {
                format = '0';
            } else if (valType === 'number' && (header === 'Payout' || header === 'Margin' || header === 'Revenue')) {
                format = '$0.00';
            } else if (valType === 'object' && header === 'Date') {
                format = 'yyyy-mm-dd';
            } else {
                format = '@';
            }

            const formatRange = this.sheet.getRange(colLetter + ":" + colLetter);
            formatRange.setNumberFormat(format);
        }
    }
}
