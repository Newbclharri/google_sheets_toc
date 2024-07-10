

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

        let newContentIdsToAddToTocSheet

        try {
            //logic
            //currentContentIds
            const currContentIds = this.myToc.fetchSheetIdsNotEqualTo(this.sheetId);
            console.log("CURRENT INCLUDES 0: ", currContentIds, currContentIds.includes(0));
            currContentIds.forEach(id => console.log(`in currIds: ${id} : ${typeof id}`,))
            

            //previousContentIds
            const storedContentIds = this.myToc.getStoredContentIds() || this.myToc.getContentIds;
            console.log("INITIAL: ", storedContentIds);

            // Get a list of content Ids already on the TOC sheet
            const contentIdsOnTocSheet = this.getContentIdsFromTocSheet();
            console.log(`IDS ON SHEET: ${contentIdsOnTocSheet}`)
            contentIdsOnTocSheet.forEach(id => console.log(`ID TYPES IN contenteIdsOnSheet: ${id}: TYPE: ${typeof id}`))
            console.log("INCLUDES 0: ",contentIdsOnTocSheet.includes(0))

            //insertedTabs
            newContentIdsToAddToTocSheet = this.findDifferences(currContentIds, contentIdsOnTocSheet);
            newContentIdsToAddToTocSheet.forEach(id => console.log(`ID TYPES IN newContentIds: ${id}: TYPE: ${typeof id}`))

            const uniqueAndValidIds = newContentIdsToAddToTocSheet.filter(id =>{

                console.log(`ID ${id} is included on the toc sheet: `, contentIdsOnTocSheet.includes(id))
                console.log(`ID: ${id} ${typeof id}`)
             return this.myToc.isValidSheetId(id) && !contentIdsOnTocSheet.includes(id)
            });

            // Filter already stored ids
            console.log("UNIQUE CONTENT IDS: ", uniqueAndValidIds);

            removeTocSheetIdFromInsertedIds(uniqueAndValidIds, this.sheetId);

            if (uniqueAndValidIds && uniqueAndValidIds.length) {

                // Create sheet links
                const links = this.myToc.createSheetLinks(uniqueAndValidIds);

                // Check if links were created successfully
                if (!links.length) {
                    console.warn('No links were created.');
                }

                // Get the named range
                const range = this.myToc.getRangeContents();
                if (!range) {
                    throw new Error('Could not get contents range.');
                }

                console.log("RANGE CONTENTS START ROW: ", range.getRow());
                const rangeStartRow = range.getRow();

                // Get the sheet and range details
                const sheet = range.getSheet();
                const rangeStartColumn = range.getColumn();

                // Shift cells down to insert new sheet links at the top of the range            
                const rangetoInsertCells = sheet.getRange(rangeStartRow, rangeStartColumn, links.length, sheet.getLastColumn());
                this.shiftCellsDown(rangetoInsertCells);

                // Define the range to paste the links
                const headerRow = this.myToc.getRangeHeader().getLastRow();
                const rangeToPaste = sheet.getRange(headerRow + 1, rangeStartColumn, links.length, 1);
                rangeToPaste.setRichTextValues(links);

                // Call additional functions if necessary
                this.myToc.setSheetDataByIdPropertiesFromLinks(links)
                // updateNamedRange();
                const newRange = sheet.getRange(rangeStartRow, rangeStartColumn, range.getNumRows() + rangeToPaste.getNumRows());
                this.updateNamedRange(newRange);

                // updateContentIds();
                this.myToc.setContentIds(currContentIds);

                // updateTitles();
                this.myToc.updateTitlesByIds(uniqueAndValidIds);
            }


            //save TOC state
            //this.myToc.save();
            // this.myToc.saveBackup();

        } catch (err) {
            console.error('Error processing inserted content:', err.stack);

            // if(uniqueAndValidIds && uniqueAndValidIds.length){
            //     uniqueAndValidIds.forEach(id =>{ 
            //         this.myToc.addSheetDataById(id);
            //     })
            // }
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
            //Sheet Ids currently on TOC sheet
            const contentIdsOnTocSheet = this.getContentIdsFromTocSheet();
            console.log("CONTENT IDS ON TOC SHEET: ", contentIdsOnTocSheet)

            // Get current list of sheet IDs
            //const currSheetIds = this.myToc.fetchSheetIds(); // Returned data is type number


            // Fetch stored sheet data
            const storedSheetData = this.myToc.getSheetDataById();

            // Convert stored sheet IDs from string to number
            //const storedSheetIds = Object.keys(storedSheetData).map(ele => Number(ele));

            // Find differences between current and stored sheet IDs
            const contentsToDeleteByIds = contentIdsOnTocSheet.filter(id => {
                console.log(id, "IS INVALID: ", !this.myToc.isValidSheetId(id));
                return !this.myToc.isValidSheetId(id)
            });
            contentsToDeleteByIds
            console.log(`CONTENT IDS TO DELETE: ${contentsToDeleteByIds}`);

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
                const linkId = linkUrl ? this.myToc.getSheetGIDFromRichTextUrl(linkUrl) : null;

                // Check for deleted sheet ID or name

                const sheetToDeleteId = contentsToDeleteByIds.find(id => {


                    return id === linkId || storedSheetData[id].name === linkName

                });

                // Remove deleted sheet ID from stored sheet data object
                //delete storedSheetData[deletedSheetId];

                if (sheetToDeleteId !== undefined && !this.myToc.isValidSheetId(sheetToDeleteId)) {
                    rows.push(startRow + rowIndex);

                    // Remeve deleted sheet id from array processing
                    const index = contentsToDeleteByIds.indexOf(sheetToDeleteId);
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
        let rangeContents;
        try {
            // Get TOC contents
            rangeContents = this.myToc.getRangeContents();

            // Get IDs from content sheet links
            const idsOnTocSheet = rangeContents.getRichTextValues()
                // Flatten the 2D array
                .flatMap(row => row)
                // Map ID from link to the 1D array
                .map(link => {
                    const url = link.getLinkUrl();
                    if (url) {
                        const id = this.myToc.getSheetGIDFromRichTextUrl(url);
                        return !isNaN(id) ? id : null;
                    }
                    return null;
                })
                // Filter out null values
                .filter(id => id !== null);

            return idsOnTocSheet;

        } catch (err) {
            console.error("Error in getContentIdsFromTocSheet:", err.stack);
        }
    }


    findDifferences(arry1, arry2) {
        return arry1
            .filter(element => !arry2.includes(element))
            .concat(arry2.filter(element => !arry1.includes(element)));
    }
    updateNamedRange(newRange) {
        try {
            // Set named range
            this.myToc.setNamedRange(this.myToc.rangeContentsName, newRange)
            // Update range a1notation
            const a1Notation = newRange.getA1Notation();
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