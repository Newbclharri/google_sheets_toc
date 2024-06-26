class NamedRangeHandler{
    constructor(name, range, obj){
        this.name = name;
        this.range = range;
        this.spreadsheetUtil = SpreadsheetUtility.getInstance(); // SpreadsheetApp
        this.spreadsheet = this.spreadsheetUtil.getActive();
        this.activeSheet = this.spreadsheet.getActiveSheet();
        this.rangeHeaderName = obj.rangeHeaderName;
        this.rangeContentsName = obj.rangeContentsName;
        this.sheetId = obj.sheetId;
    }

    handleRangeEdit(e){    
        if(this.range && this.activeSheet.getSheetId() === this.sheetId){
            console.log("tocSheet edit detected.")
            const rangeContents = this.spreadsheetUtil.getActive().getRangeByName(this.rangeContentsName)
            if(this.range.columnStart === rangeContents.getColumn()){
                if(this.range.rowStart <= rangeContents.getEndRow()){
                    console.log("target",e.value, e.oldValue)
                }
            }       

        }

    }

}