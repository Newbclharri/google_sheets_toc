class SheetChangeHandler {
    constructor(obj = {}, id) {
        this.sheet = obj;
        this.backupKey = obj.backupKey;
        this.targetSheetId = id //obj.sheetId;
        this.spreadsheetUtil = SpreadsheetUtility.getInstance(); //|| SpreadsheetApp;
        this.activeSheet = this.spreadsheetUtil.getActive().getActiveSheet();
        this.uI = new UiUtil();
        this.rangeHeaderName = obj.rangeHeaderName;
        this.rangeContentsName = obj.rangeContentsName;
    }


    handleChange(e) {
        const activeSheetId = this.activeSheet.getSheetId()
        console.log("active id: ", activeSheetId, " | targetId", this.targetSheetId)
        if (activeSheetId === this.targetSheetId) {
            console.log("tocSheet change detected.")
            //EDIT, REMOVE_COLUMN, INSERT_COLUMN, REMOVE_ROW, INSERT_ROW all require sheet data range values backup
            this.updateDataRangeValues();
            switch (e.changeType) {
                case "REMOVE_COLUMN":
                case "INSERT_COLUMN":
                case "REMOVE_ROW":
                case "INSERT_ROW":
                    this.updateRanges() //above cases may change the TOC named ranges (header range and contents range)
                    break;
                case "OTHER": //tab is renamed
                    this.handleRename();
                    break;

                default:
                    console.log("SheetChangeHandler.js other changeTypes(FORMAT, EDIT): ", e.changeType)
            }
            /////////////SAVE UPDATED PROPERTIES//////////////
            this.sheet.save();
            this.sheet.saveBackup();
        }
    }


    handleRename() {
        const newName = this.activeSheet.getName();
        if (newName !== this.sheet.name) {
            this.sheet.setName(newName);
        }
    }

    updateRanges() {
        const newRangeHeaderA1Notation = this.spreadsheetUtil
            .getActive()
            .getRangeByName(this.rangeHeaderName)
            .getA1Notation();

        const newRangeContentsA1Notation = this.spreadsheetUtil
            .getActive()
            .getRangeByName(this.rangeContentsName)
            .getA1Notation();
        this.sheet.updateState({
            rangeHeaderA1Notation: newRangeHeaderA1Notation,
            rangeContentsA1Notation: newRangeContentsA1Notation
        });      
       
    }

    updateDataRangeValues() {
        const values = this.activeSheet.getDataRange().getValues();
        if(values){
            this.sheet.updateState({ dataRangeValues: values });

            // this.dataRangeValues = values;
            this.sheet.updateBackup()
        }
    }

    getDataRangeValues() {
        return this.values
    }

    getSheetUpdates() {
        const updates = this.sheet.toJSON();

        return updates;
    }

    // getBackupData() {
    //     const data = this.getDataRangeValues();
    //     const backupData = { ...this.getSheetUpdates(), data }
    //     return backupData;
    // }
}



