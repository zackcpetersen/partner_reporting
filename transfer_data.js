class TransferApiData {
    constructor() {
        this.querySpreadsheet = SpreadsheetApp.openById('1jPFeGJgeU0nyHZx9DXTL0wagIQM3Cbl5AyzD37ky3ck');
        this.querySheet = this.querySpreadsheet.getSheetByName('Query Data');
        this.queryData = this.querySheet.getDataRange().getValues();
        this.metricsSpreadsheet = SpreadsheetApp.openById('1xjeR7V4Qm5OCHW2iWo1zjbVriIX1NsmMUzK1sLNhK1g');
        this.webSheet = this.metricsSpreadsheet.getSheetByName('Web Funnel');
        this.callSheet = this.metricsSpreadsheet.getSheetByName('Call Center Funnel');
        this.queryValues = this.getQueryData();
    }
    transferQueryData() {
        this.transferData(this.callSheet);
        this.transferData(this.webSheet);
    }
    getQueryData() {
        let queryValues = {
            'Leads': {},
            'Revenue': {}
        };
        // loop through query data and add each row to an object
        let lastCol = columnToLetter(this.queryData[0].length);
        let clientLength = this.queryData.length / 2;
        for (let i = 0; i < this.queryData.length; i++) {
            let client = this.queryData[i][0];
            let rowIndex = i + 1;
            let row = this.querySheet.getRange('B' + rowIndex + ':' + lastCol + rowIndex);
            if (client && client !== 'Client' && (rowIndex < clientLength + 1)) {
                queryValues['Leads'][client] = row.getValues();
            } else if (client && client !== 'Client' && (rowIndex > clientLength + 1) && client !== 'AcademixDirectCall') {
                queryValues['Revenue'][client] = row.getValues();
            }
        }
        return queryValues;
    }
    transferData(sheet) {
        // loop through column A in web and call sheets, paste matching values from queryValues
        let sheetData = sheet.getDataRange().getValues();
        let lastCol = columnToLetter(sheetData[0].length);
        for (let i = 0; i < sheetData.length; i++) {
            let client = sheetData[i][0];
            let rowIndex = i + 1;
            let range = sheet.getRange('B' + rowIndex + ':' + lastCol + rowIndex);
            if (client.includes('Revenue')) {
                if (this.queryValues['Revenue'][client.split(" - ")[0]]) {
                    range.setValues(this.queryValues['Revenue'][client.split(" - ")[0]]);
                }
            } else {
                if (this.queryValues['Leads'][client]) {
                    range.setValues(this.queryValues['Leads'][client]);
                }
            }
        }
    }
}


function transferQueryData() {
    const apiTransfer = new TransferApiData();
    apiTransfer.transferQueryData();
}
