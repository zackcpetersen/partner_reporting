class QuerySheet {
    constructor() {
        // =IFNA(QUERY('API Overview - For Data Studio'!$A:$P, "SELECT SUM(A) WHERE (B = date '"&text(datevalue(DA$1), "yyyy-mm-dd")&"' AND M = '"&$A9&"') label sum(A) ''"),0)
        this.date = new Date()
        this.ss = SpreadsheetApp.openById('1jPFeGJgeU0nyHZx9DXTL0wagIQM3Cbl5AyzD37ky3ck');
        this.querySheet = this.ss.getSheetByName('Query Data');
        this.dataSheet = this.ss.getSheetByName('API Overview - For Data Studio');
        this.queryData = this.querySheet.getDataRange().getValues();
        this.yesterdayCol = this.findYesterdayCol();
        this.leadsStartingRow = 2;
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
        const cell = this.querySheet.getRange(columnToLetter(this.yesterdayCol) + startingRow);
        cell.setValue(this.buildQuery(sumCol, clientCell));

        const destination = this.querySheet.getRange(startingRow, cell.getColumn(), numRows, 1);
        cell.copyTo(destination);
    }
    findYesterdayCol() {
        for (let i = 0; i < this.queryData[0].length; i++) {
            if (typeof this.queryData[0][i] === 'object'){
                if (this.formattedDate(this.queryData[0][i]) === this.yesterday()) {
                    return i + 1;
                }
            }
        }
    }
    buildQuery(sumCol, clientCell) {
        const lastCol = this.dataSheet.getDataRange().getValues()[0].length;
        const dataRange = '$A:' + '$' + columnToLetter(lastCol);
        const queryDataRange = "'" + this.dataSheet.getName() + "'" + "!" + dataRange;
        const dateCell = columnToLetter(this.yesterdayCol) + "$" + 1;
        const leadsQuery = "SELECT SUM(" + sumCol + ") WHERE (B = date '\"&text(datevalue(" + dateCell + "), \"yyyy-mm-dd\")&\"' AND O = '\"&" + clientCell + "&\"') label sum(" + sumCol + ") ''";

        return "=IFNA(QUERY(" + queryDataRange + ", " + '"' + leadsQuery + '"' + "), 0)";
    }
    yesterday() {
        const oneDayBack = this.date.getUTCDate() - 1;
        return this.date.getUTCFullYear() + "-" + this.date.getUTCMonth() + "-" + oneDayBack;
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
