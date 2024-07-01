class NamedRangeHandler {
    constructor(name, range, obj = {}, id) {
        this.name = name;
        this.range = range;
        this.myToc = obj;
        this.sheetId = id //obj.sheetId;
        this.ssUtil = SpreadsheetUtility.getInstance(); // SpreadsheetApp
        this.spreadsheet = this.ssUtil.getActive();
        this.activeSheet = this.spreadsheet.getActiveSheet();
        this.rangeHeaderName = obj.rangeHeaderName;
        this.rangeContentsName = obj.rangeContentsName;
    }

    handleRangeEdit(e) {
        // console.log("inside NamedRangeHandler: this.range: ", this.range)
        //possible try catch...
        if (this.wasEditedContentRange(this.myToc, this.activeSheet)) {
            if (this.range.rowStart <= rangeContents.getEndRow()) {
                console.log("target", e.value, e.oldValue)
            }
        }
    }

    wasEditedContentRange(myToc, sheet) {
        try {
            const currentColumn = sheet.getActiveRange().getColumn();
            const rangeContents = myToc.getRangeByName(myToc.rangeContentsName);
            console.log("CURRENT COL: ", currentColumn)

            if (rangeContents) {
                return currentColumn === rangeContents.getColumn();
            }
        }catch(err){
            console.error("Could not get named range of contents: ", err.stack)
        }
    }
}