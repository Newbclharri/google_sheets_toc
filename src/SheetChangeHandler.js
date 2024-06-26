class SheetChangeHandler {
    constructor({ name, sheetId, key, backupKey, rangeHeaderName, rangeContentsName }) {
        this.name = name;
        this.key = key;
        this.backupKey = backupKey;
        this.targetSheetId = sheetId;
        this.spreadsheetUtil = SpreadsheetUtility.getInstance(); //|| SpreadsheetApp;
        this.activeSheet = this.spreadsheetUtil.getActive().getActiveSheet();
        this.uI = new UiUtil();
        this.rangeHeaderName = rangeHeaderName;
        this.rangeContentsName = rangeContentsName;
    }


    handleChange(e) {
        const activeSheetId = this.activeSheet.getSheetId()
        if (activeSheetId === this.targetSheetId) {
            console.log("tocSheet change detected.")
            //EDIT, REMOVE_COLUMN, INSERT_COLUMN, REMOVE_ROW, INSERT_ROW all require sheet range values backup
            this.updateData();
            switch (e.changeType) {
                case "REMOVE_COLUMN":
                case "INSERT_COLUMN":
                case "REMOVE_ROW":
                case "INSERT_ROW":
                    this.updateRanges() //above cases may change the TOC named ranges (header range and contents range)
                    break;
                case "OTHER": //tab is renamed
                    this.name = this.activeSheet.getName();
                    break;
    
                default:
                    console.log("other changeTypes(FORMAT, EDIT): ", e.changeType)
            }
        }
    }
    handleRename(name) {
        this.name = this.activeSheet.getName();
    }

    updateRanges() {
        const rangeHeaderA1Notation = this.spreadsheetUtil
            .getActive()
            .getRangeByName(this.rangeHeaderName)
            .getA1Notation();

        const rangeContentsA1Notation = this.spreadsheetUtil
            .getActive()
            .getRangeByName(this.rangeContentsName)
            .getA1Notation();
        this.rangeHeaderA1Notation = rangeHeaderA1Notation;
        this.rangeContentsA1Notation = rangeContentsA1Notation;
        console.log("New Range Header: ", this.rangeHeaderA1Notation, " New range contents: ", this.rangeContentsA1Notation);
    }

    updateData() {
        const values = this.activeSheet.getDataRange().getValues();
        this.values = values

    }

    getDataRangeValues() {
        return this.values
    }

    getSheetUpdates() {
        const {name, sheetId, key, backupKey, rangeHeaderName, rangeHeader, rangeContentsName, rangeContents} = this;
        
        return { name, key, backupKey, sheetId, rangeHeaderName, rangeContentsName, rangeHeader, rangeContents };
    }

    getBackupData(){
        const data = this.getDataRangeValues();
        const backupData = {...this.getSheetUpdates(), data}
        return backupData;
    }
}



