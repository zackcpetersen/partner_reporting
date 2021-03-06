class QuerySheet {
    constructor() {
        // =IFNA(QUERY('API Overview - For Data Studio'!$A:$P, "SELECT SUM(A) WHERE (B = date '"&text(datevalue(DA$1), "yyyy-mm-dd")&"' AND M = '"&$A9&"') label sum(A) ''"),0)
        this.ss = SpreadsheetApp.openById('1jPFeGJgeU0nyHZx9DXTL0wagIQM3Cbl5AyzD37ky3ck');
        this.querySheet = this.ss.getSheetByName('Query Data');
        this.dataSheet = this.ss.getSheetByName('API Overview - For Data Studio');
        this.queryData = this.querySheet.getDataRange().getValues();
        this.queryDays = 30;  // number of days to run queries
        this.daysBackCol = this.findDaysBackCol();  // dynamically gets today minus this.queryDays column letter
        this.leadsStartingRow = 2;  // row after the header row
        this.clientLength = (this.queryData.length / 2) - 1;  // for blank row and header
        this.revenueStartingRow = this.leadsStartingRow + this.clientLength + 1;  // for blank row
    }
    fillQueries() {
        // copy cell function down for lead numbers
        let clientCell = '$A' + this.leadsStartingRow;
        this.setQueryValues(this.leadsStartingRow, clientCell, 'A', this.clientLength);

        // copy cell function down for revenue
        clientCell = '$A' + this.revenueStartingRow;
        this.setQueryValues(this.revenueStartingRow, clientCell, 'D', this.clientLength);

        // copy and paste as values only
        this.removeFunctions();
    }
    setQueryValues(startingRow, clientCell, sumCol, numRows) {
        const cell = this.querySheet.getRange(columnToLetter(this.daysBackCol) + startingRow);
        cell.setValue(this.buildQuery(sumCol, clientCell));

        const destination = this.querySheet.getRange(startingRow, cell.getColumn(), numRows, this.queryDays);
        cell.copyTo(destination);
    }
    findDaysBackCol() {
        const daysBack = this.daysBack()
        // Could potentially accomplish this with a hash table, but I still need the index so
        // I'm not sure how much faster it would actually be
        for (let i = 0; i < this.queryData[0].length; i++) {
            if (typeof this.queryData[0][i] === 'object'){
                if (this.formattedDate(this.queryData[0][i]) === daysBack) {
                    return i + 1;
                }
            }
        }
    }
    buildQuery(sumCol, clientCell) {
        const lastCol = this.dataSheet.getDataRange().getValues()[0].length;
        const dataRange = '$A:' + '$' + columnToLetter(lastCol);
        const queryDataRange = "'" + this.dataSheet.getName() + "'" + "!" + dataRange;
        const dateCell = columnToLetter(this.daysBackCol) + "$" + 1;
        const leadsQuery = "SELECT SUM(" + sumCol + ") WHERE (B = date '\"&text(datevalue(" + dateCell + "), \"yyyy-mm-dd\")&\"' AND O = '\"&" + clientCell + "&\"') label sum(" + sumCol + ") ''";

        return "=IFNA(QUERY(" + queryDataRange + ", " + '"' + leadsQuery + '"' + "), 0)";
    }
    daysBack() {
        const daysBack = new Date().setDate(new Date().getDate() - this.queryDays)
        const daysBackDate = new Date(daysBack)
        return this.formattedDate(daysBackDate)
    }
    formattedDate(date) {
        return date.getUTCFullYear() + "-" + date.getUTCMonth() + "-" + date.getUTCDate();
    }
    removeFunctions() {
        this.querySheet.getDataRange().setValues(this.querySheet.getDataRange().getValues());
    }
}


function fillQueries() {
    const queries = new QuerySheet();
    queries.fillQueries();
}
