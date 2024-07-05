

class GridChangeHandler {
    constructor(obj = {}, id) {
        this.myToc = obj;
        this.sheetId = id
        this.ssUtil = SpreadsheetUtility.getInstance();
        // this.uI = UiUtil.getInstance(); //SpreadsheetApp.getUi();
        this.currentListOfSheetIds = this.myToc.fetchSheetIds();
        this.propsStorage = PropertiesServiceStorage.getInstance(); //PropertiesService
        //this.handleRemoveGrid();

    }

    handleRemoveGrid() {
        if (this.isRemovedTocTab()) {
            this.handleUserRemovesTocTab();
        } else {
            this.handleUserDeletesSheet();
        }
    }


    isRemovedTocTab() {
        //logic here
        return !this.myToc.fetchSheetIds().some(id => id === this.sheetId)
    }

    handleUserInsertsSheet() {
        const activeSheetId = this.ssUtil.getActive().getActiveSheet().getSheetId();

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
            removeTocSheetIdFromInsertedIds(insertedContentIds, this.sheetId);

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
            const rangetoInsertCells = sheet.getRange(range.getRow(), rangeStartColumn, links.length, 1);
            this.shiftCellsDown(rangetoInsertCells);

            // Define the range to paste the links
            const rangeToPaste = sheet.getRange(2, rangeStartColumn, links.length, 1);
            rangeToPaste.setRichTextValues(links);

            // Call additional functions if necessary
            this.myToc.setSheetDataByIdPropertiesFromLinks(links)
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

        function removeTocSheetIdFromInsertedIds(insertedIds, targetId) {
            return removeElementFromArray(insertedIds, targetId);

        }

        function removeElementFromArray(array, element) {
            const targetIndex = array.indexOf(element);
            if (targetIndex > -1) {
                array.splice(targetIndex, 1)
            }
            return array;
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

    

    handleUserDeletesSheet() {
        try {
            // Get current list of sheet IDs
            const currSheetIds = this.myToc.fetchSheetIds(); // Returned data is type number
    
            // Fetch stored sheet data
            const storedSheetData = this.myToc.getSheetDataById();
    
            // Convert stored sheet IDs from string to number
            const storedSheetIds = Object.keys(storedSheetData).map(Number);
    
            // Find differences between current and stored sheet IDs
            const contentsToDeleteByIds = this.findDifferences(currSheetIds, storedSheetIds);
            console.log(`DIFFERENCES BEFORE: ${contentsToDeleteByIds}`);
    
            // Get the range of content
            const range = this.myToc.getRangeContents();
    
            if (!range) {
                console.error("Could not get the range to remove contents.");
                return;
            }
    
            const startRow = range.getRow();
            const sheet = range.getSheet();
            const contentNames = range.getValues().map(row => row[0]);
            const contentLinks = range.getRichTextValues().map(row => row[0]);
    
            // Accumulate rows to delete using .reduce
            const rowsToDelete = contentLinks.reduce((rows, link, rowIndex) => {
                const linkUrl = link.getLinkUrl();
                const linkName = link.getText();
                const linkId = linkUrl ? this.myToc.getSheetGIDFromRichText(linkUrl) : null;
    
                // Check for deleted sheet ID or name
                const deletedSheetId = contentsToDeleteByIds.find(id => id === linkId || storedSheetData[id].name === linkName);

                // Remove deleted sheet ID from stored sheet data object
                delete storedSheetData[deletedSheetId];
                
                if (deletedSheetId !== undefined) {
                    rows.push(startRow + rowIndex);
                    
                    // Remeve deleted sheet id from array processing
                    const index = contentsToDeleteByIds.indexOf(deletedSheetId);
                    if (index > -1) {
                        contentsToDeleteByIds.splice(index, 1);
                    }
                }
                return rows;
            }, []);
            
            // Delete sheet links from bottom of range to the top most row:
            // Avoids index shifting as rows are deleted
            for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                sheet.deleteRow(rowsToDelete[i]);
            }
    
        } catch (error) {
            console.error('Error handling user removes content tab:', error.stack);
        }
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
            console.error("Error in getContentIdsFromTocSheet:", err.stack);
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
            console.log("NEW LAST ROW: ", range.getLastRow())
            const a1Notation = range.getA1Notation();
            this.myToc.rangeContentsA1Notation = a1Notation;
            console.log(`New range updated successfully: ${a1Notation}`);
        } catch (err) {
            console.error(err.stack);
        }
    }

    removeBlanksFromRange(range) {
        if (range) {
            const sheet = range.getSheet();
            const rangeRowStart = range.getRow();
            const arrayEmptyIndices = [];
            const values = range.getValues();
            let blankRow;
            //search range values for blank cells and push the index into an array
            values.forEach((row, blankRowIndex) => {
                row.forEach(value => {
                    if (value == "") {
                        blankRow = rangeRowStart + blankRowIndex;
                        sheet.deleteRow(blankRow);
                    }
                })
            });
        }
    }

    shiftCellsDown(range) {
        const direction = this.ssUtil.spreadsheetApp.Dimension.ROWS;
        range.insertCells(direction);
    }
}