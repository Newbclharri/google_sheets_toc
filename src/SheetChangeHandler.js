class SheetChangeHandler {
    constructor(e, obj = {}, id, activeRange) {
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
        this.rangeContents = obj.getRangeContents();
        this.headerRow = obj.getRangeHeader().getLastRow() || 1;
        this.targetRangeRowStart = this.rangeContents.getRow() || 2;
        this.targetRangeColumnStart = this.rangeContents.getColumn();
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

                    if (this.wasShiftedDownContents()) {
                        this.removeInvalidRowsAboveNamedRange()
                    }

                    // if (this.wasChangedContentRange()) {
                    //     const activeRowStart = this.activeRange.getRow();
                    //     const rangContentsRowStart = this.myToc.getRangeContents().getRow();
                    //     if(activeRowStart < rangContentsRowStart){
                    //         this.removeInvalidRowsAboveNamedRange();
                    //     }else{
                    //         this.removeInvalidRowsFromRange();
                    //     }
                    // }
                    // this.updateContentRange(this.eventObject) //above cases may change the TOC named ranges (header range and contents range)
                    break;
                case "OTHER": //tab is renamed
                    //this.handleRename();
                    //this.removeInvalidRowsFromRange();
                    if(this.wasShiftedDownContents()){
                        this.removeInvalidRowsAboveNamedRange();
                    }

                    break;

                default:
                    console.log("SheetChangeHandler.js other changeTypes(FORMAT, EDIT): ", this.changeType)
            }

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

    getActiveSheetId() {
        return this.activeRange.getSheet().getSheetId();
    }

    wasChangedHeaderRow() {
        try {
            const activeSheetId = this.getActiveSheetId();
            if (activeSheetId == this.targetSheetId) {
                return this.activeRange.getRow() === this.headerRow;
            }
        } catch (err) {
            console.error("Error attempting to retrieve the header range", err.stack);
            return false;
        }
    }

    /**
   * Detects a user change to the contents range
   * in the TOC sheet
   * @returns {Boolean}
   */
    wasChangedContentRange() {
        try {
            // Get the active sheet ID to localize user changes
            const activeSheetId = this.getActiveSheetId();

            // Determine if the change occurred in the TOC sheet
            if (activeSheetId === this.targetSheetId) {
                // Get the start and end rows of the active range
                const startRow = this.activeRange.getRow();
                const endRow = this.activeRange.getLastRow();

                // Get the contents range to compare with the active range
                const rangeContents = this.myToc.getRangeContents();
                const rangeContentsEndRow = rangeContents.getLastRow();
                console.log(`CONTENTS END ROW: ${rangeContentsEndRow}`);

                // Check if the active range is within the contents range (excluding the header row)
                if (startRow > this.headerRow && endRow <= rangeContentsEndRow) {
                    return true;
                }
            }

            // Return false if the change did not occur in the specified range
            return false;

        } catch (err) {
            // Log an error message if an exception occurs
            console.error("Error retrieving the contents range:", err);
            return false; // Ensure a boolean is always returned
        }
    }

    wasShiftedDownContents() {
        if (this.targetRangeRowStart > this.headerRow + 1) {
            return true;
        }

        return false;

    }

    removeInvalidRowsFromRange() {
        const range = this.myToc.getRangeContents();
        this.myToc.removeInvalidRowsFromRange(range);
    }

    removeInvalidRowsAboveNamedRange() {
        const headerRow = this.headerRow
        const sheet = this.rangeContents.getSheet();
        const rangeToClean = sheet.getRange(headerRow + 1, this.targetRangeColumnStart, this.targetRangeRowStart);
        this.myToc.removeInvalidRowsFromRange(rangeToClean);
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



