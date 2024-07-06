class SheetChangeHandler {
    constructor(e, obj = {}, id, activeRange, headerRow, targetRangeRowStart) {
        this.eventObject = e;
        this.changeType = e.changeType
        this.myToc = obj;
        this.backupKey = obj.backupKey;
        this.targetSheetId = id //obj.sheetId;
        this.spreadsheetUtil = SpreadsheetUtility.getInstance(); //|| SpreadsheetApp;
        this.activeRange = activeRange;
        this.activeSheet = this.spreadsheetUtil.getActive().getActiveSheet();
        this.uI = new UiUtil();
        this.rangeHeaderName = obj.rangeHeaderName;
        this.rangeContentsName = obj.rangeContentsName;
        this.headerRow = headerRow || 1;
        this.targetRangeRowStart = targetRangeRowStart || 2;
    }


    handleChange(changeType = this.changeType, activeRange = this.activeRange) {
        const activeSheetId = this.activeSheet.getSheetId()
        console.log("active id: ", activeSheetId, " | targetId:", this.targetSheetId)
        if (activeSheetId === this.targetSheetId) {
            console.log("tocSheet change detected.")
            //EDIT, REMOVE_COLUMN, INSERT_COLUMN, REMOVE_ROW, INSERT_ROW all require sheet data range values backup
            this.updateDataRangeValues();
            switch (changeType) {
                case "REMOVE_COLUMN":
                case "INSERT_COLUMN":
                case "REMOVE_ROW":
                case "INSERT_ROW":
                    console.log("ACTIVE RANGE: Row: ", activeRange.getRow())
                    if (this.wasChangedHeaderRow()) {
                        console.log("USER CHANGED HEADER ROW")
                        this.removeExcessHeaderRows();
                    }
                    this.updateContentRange(this.eventObject) //above cases may change the TOC named ranges (header range and contents range)
                    break;
                case "OTHER": //tab is renamed
                    this.handleRename();
                    break;

                default:
                    console.log("SheetChangeHandler.js other changeTypes(FORMAT, EDIT): ", this.changeType)
            }
            /////////////SAVE UPDATED PROPERTIES//////////////
            this.myToc.save();
            this.myToc.saveBackup();
        }
    }


    handleRename() {
        const newName = this.activeSheet.getName();
        if (newName !== this.myToc.name) {
            this.myToc.setName(newName);
        }
    }

    updateContentRange() {
        try {
            const range = this.myToc.getRangeContents();
            const newRangeContentsA1Notation = range.getA1Notation();
            this.myToc.updateState({
                rangeContentsA1Notation: newRangeContentsA1Notation
            });
        } catch (err) {
            console.error("Could not update range: ", err)
            console.log(err.stack);
        }
    }

    updateDataRangeValues() {
        const values = this.activeSheet.getDataRange().getValues();
        if (values) {
            this.myToc.updateState({ dataRangeValues: values });

            // this.dataRangeValues = values;
            this.myToc.updateBackup()
        }
    }

    getDataRangeValues() {
        return this.values
    }

    getSheetUpdates() {
        const updates = this.myToc.toJSON();

        return updates;
    }

    wasChangedHeaderRow() {
        try {
            const activeSheetId = this.activeRange.getSheet().getSheetId();
            if(activeSheetId == this.targetSheetId){
                return this.activeRange.getRow() === this.headerRow;
            }
        } catch (err) {
            console.error("Error attempting to retrieve the header range", err.stack);
        }
    }

    removeExcessHeaderRows() {
        try {
            const sheet = this.activeRange.getSheet();
            const startRow = this.activeRange.getRow();
            const numRows = this.activeRange.getNumRows();
            sheet.deleteRows(startRow, numRows);
        } catch (err) {
            console.error("Error handling the header range. ", err.stack);
        }
    }


}



