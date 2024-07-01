

class GridChangeHandler {
    constructor(obj = {}, id) {
        this.myToc = obj;
        this.sheetId = id
        this.ssUtil = SpreadsheetUtility.getInstance();
        this.uI = UiUtil.getInstance(); //SpreadsheetApp.getUi();
        this.currentListOfSheetIds = this.myToc.fetchSheetIds();
        this.propsStorage = PropertiesServiceStorage.getInstance(); //PropertiesService
        //this.handleRemoveGrid();

    }

    handleRemoveGrid() {
        if (this.isRemovedTocTab()) {
            this.handleUserRemovesTocTab();
        } else {

        }
    }


    isRemovedTocTab() {
        //logic here
        return !this.myToc.fetchSheetIds().some(id => id === this.sheetId)
    }

    handleUserInsertsSheet() {
        try {
            //logic
            //currentContentIds
            const currContentIds = this.myToc.fetchSheetIdsNotEqualTo(this.sheetId);
            console.log("CURRENT: ", currContentIds)
            //previousContentIds
            const initialContentIds = this.getContentIdsFromTocSheet() || this.myToc.getContentIds;
            console.log("INITIAL: ", initialContentIds)
            //insertedTabs
            const insertedContentIds = this.findDifferences(currContentIds, initialContentIds);
            console.log("INSERTED CONTENT IDS: ", insertedContentIds)
            console.log()

            // Check if insertedContentIds exists and has length
            if (!insertedContentIds.length) {
                throw new Error('No content IDs to process.');
            }

            // Create sheet links
            const links = this.myToc.createSheetLinks(insertedContentIds);

            // Check if links were created successfully
            if (!links.length) {
                throw new Error('No links were created.');
            }

            // Get the named range
            const range = this.myToc.getRangeContents();
            if (!range) {
                throw new Error('Could not get contents range.');
            }

            // Get the sheet and range details
            const sheet = range.getSheet();
            const rangeStartColumn = range.getColumn();

            //sheft cells down to insert new sheet links at the top of the range            
            const rangetoInsertCells = sheet.getRange(range.getRow(),rangeStartColumn,links.length,1);
            this.shiftCellsDown(rangetoInsertCells);
            
            // Define the range to paste the links
            const rangeToPaste = sheet.getRange(2, rangeStartColumn, links.length, 1);
            rangeToPaste.setRichTextValues(links);
            
            // Call additional functions if necessary
            // updateNamedRangeRows();
            this.updateNamedRangeRows()

            // updateContentIds();
            this.myToc.setContentIds(currContentIds);

            // updateTitles();
            this.myToc.updateTitlesByIds(insertedContentIds);

            //save TOC state
            //this.myToc.save();
           // this.myToc.saveBackup();

        } catch (err) {
            console.error('Error processing inserted content:', err.stack);
            return; // Early return on error
        }

    }



    handleUserRemovesTocTab() {
        // const key = this.myToc.backupKey;
        const backup = this.myToc.toJSON();
        //logic
        //update to potentially re-add links
        this.myToc.updateState({ "allSheetIds": this.currentListOfSheetIds })
        this.myToc.restore(backup);
        this.uI.alert("Select 'Remove' from Table of Contents menu to remove this sheet.")

    }

    handleUserRemovesContentTab() {
        //logic
    }


    getContentIdsFromTocSheet() {
        //user could potentially change the range name
        let rangeContents;
        try {
            //get TOC contents
            rangeContents = this.myToc.getRangeContents();            
            const sheetNames = rangeContents.getValues().filter((row, index) => row[0] !== "").map(row => row[0])
            // console.log("SHEETNAMES: ", sheetNames)

            //Get sheetIds for each value (sheet / tab names)
            const contentSheetIds = sheetNames.map(sheetName => {
                const sheet = this.ssUtil.getSheetByName(sheetName);
                if (sheet) {
                    const id = sheet.getSheetId();
                    return !isNaN(id) ? id : null;
                }
                return null;
            }).filter(id => id !== null);;

            return contentSheetIds;
        } catch (err) {
            console.error("Error in getContentIdsFromTocSheet:",err.stack);
        }
    }

    findDifferences(arry1, arry2) {
        return arry1
            .filter(element => !arry2.includes(element))
            .concat(arry2.filter(element => !arry1.includes(element)));
    }
    updateNamedRangeRows() {
        const range = this.myToc.getRangeContents();
        if (!range) {
            console.log("Could not get the range");
            return;
        }

        try {
            //named range details
            console.log("NEW LAST ROW: " ,range.getLastRow())
            const a1Notation = range.getA1Notation();
            this.myToc.rangeContentsA1Notation = a1Notation;
            console.log(`New range updated successfully: ${a1Notation}`);
        } catch (err) {
            console.error(err.stack);
        }
    }

    removeBlanksFromRange(range){
        if(range){
            const sheet = range.getSheet();
            const rangeRowStart = range.getRow();
            const arrayEmptyIndices = [];
            const values = range.getValues();
            let blankRow;
            //search range values for blank cells and push the index into an array
            values.forEach((row, blankRowIndex) =>{
                row.forEach(value =>{
                   if(value == ""){
                    blankRow = rangeRowStart + blankRowIndex;
                    sheet.deleteRow(blankRow);
                   }
                })
            });
        }
    }

    shiftCellsDown(range){
        const direction = this.ssUtil.spreadsheetApp.Dimension.ROWS;
        range.insertCells(direction);
    }
}