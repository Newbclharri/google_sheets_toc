class NamedRangeHandler {
    constructor(name, range, obj = {}, id) {
        this.name = name;
        this.range = range;
        this.spreadsheetUtil = SpreadsheetUtility.getInstance(); // SpreadsheetApp
        this.spreadsheet = this.spreadsheetUtil.getActive();
        this.activeSheet = this.spreadsheet.getActiveSheet();
        this.rangeHeaderName = obj.rangeHeaderName;
        this.rangeContentsName = obj.rangeContentsName;
        this.sheetId = id //obj.sheetId;
    }

    handleRangeEdit(e) {
        // console.log("inside NamedRangeHandler: this.range: ", this.range)
        if (this.range && this.activeSheet.getSheetId() === this.sheetId) {
            console.log("tocSheet edit detected.")
            //possible try catch...
            const rangeContents = this.spreadsheetUtil.getActive().getRangeByName(this.rangeContentsName)
            if (this.range.columnStart === rangeContents.getColumn()) {
                if (this.range.rowStart <= rangeContents.getEndRow()) {
                    console.log("target", e.value, e.oldValue)
                }
            }

        }

    }

}