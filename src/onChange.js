function onChange(e) {
    const spreadsheetUtil = SpreadsheetUtility.getInstance();
    const propsStorage = new PropertiesServiceStorage();
    const sheetId = propsStorage.load("tocSheetId");
    let myToc;
    let tocSheetDoesExist = false;

    if (sheetId) {
        try {
            const loaded = TocSheet.load();
            myToc = new TocSheet(loaded, spreadsheetUtil, propsStorage);
            tocSheetDoesExist = myToc.doesExistSheet();
        } catch (err) {
            console.error("An error occured attempting to find the TOC sheet: ", err);
        }
    }

    if (tocSheetDoesExist) {
        console.log("SHEET EXISTS, CAN DO WORK!");
        console.log("changeType: ", e.changeType);
        onEdit(e, e.changeType);

        if (e.changeType) {
            //updateContentRange(myToc);
            switch (e.changeType) {
                case "INSERT_GRID":
                    handleGridChange(myToc, sheetId, "INSERT_GRID");
                    myToc.save();
                    myToc.saveBackup();
                    break;
                case "REMOVE_GRID":
                    handleGridChange(myToc, sheetId, "REMOVE_GRID");
                    break;
                case "OTHER":
                    // getRenamedSheetIds(spreadsheetUtil, e.changeType, myToc);
                    handleRenames(e.changeType, myToc, spreadsheetUtil);
                    // handleSheetChange(myToc, sheetId, e);
                    myToc.save();
                    myToc.saveBackup();
                    break;
                default: //"INSERT_COLUMN, REMOVE_COLUMN,  INSERT_ROW, REMOVE_ROW, OTHER, EDIT"
                    handleSheetChange(myToc, sheetId, e);
                    break;
            }
        }
    } else {
        console.log("SHEET DOES NOT EXIST. CAN'T DO WORK.");
    }
}

function handleGridChange(myToc, sheetId, changeType) {
    const gridHandler = new GridChangeHandler(myToc, sheetId);
    if (changeType === "INSERT_GRID") {
        gridHandler.handleUserInsertsSheet();
    } else if (changeType === "REMOVE_GRID") {
        gridHandler.handleRemoveGrid();
    }
}

function handleSheetChange(myToc, sheetId, e) {
    const tocChangeHandler = new SheetChangeHandler(myToc, sheetId);
    tocChangeHandler.handleChange(e);
}

function updateContentRange(myToc) {
    try {
        const range = myToc.verifyRange();
        const newRangeContentsA1Notation = range.getA1Notation();
        myToc.updateState({
            rangeContentsA1Notation: newRangeContentsA1Notation
        });
    } catch (err) {
        console.error("Could not update range: ", err)
        console.log(err.stack);
    }
}



function handleRenames(changeType, myToc, spreadsheetUtil) {
    try {
        const renamedSheetIds = getRenamedSheetIds(changeType, myToc, spreadsheetUtil);
        const rangeContents = myToc.getRangeContents();
        const storedSheetData = myToc.sheetDataById;
        //TOC sheet
        const sheet = rangeContents.getSheet();
        let newLink;

        if (renamedSheetIds.length) {
            //////////////////////////IF TOC SHEET WAS RENAMED, UPDATE NAME/////////////////

            //Find sheet data if TOC sheet was renamed
            const tocSheetId = myToc.loadTocSheetId();           
        
            if (renamedSheetIds.some(sheetData => sheetData.id == tocSheetId)) {
                const tocSheetRenameData = renamedSheetIds.find(sheetData => sheetData.id == tocSheetId)
                console.log("tocSheetRenameData: ", tocSheetRenameData)

                const targetIndex = renamedSheetIds.indexOf(tocSheetRenameData)
                //Update TOC name       
                if (targetIndex > -1) {
                    myToc.setName(tocSheetRenameData.newName)
                    console.log(myToc.name)
                    renamedSheetIds.splice(targetIndex, 1)
                }

                console.log("RENAMED SHEETIDS: BEFORE: ", renamedSheetIds)
            }

            if (!renamedSheetIds.length) {
                //If no more data to process, then exit function
                return;
            }
            /////////////////////UPDATE LINKS FOR RENAMED SHEETS//////////////////
            console.log("RENAMED SHEETIDS: AFTER: ", renamedSheetIds)

            //Get rich text values from target TOC sheet
            const links = rangeContents.getRichTextValues().map(row => row[0]);
            //Was having trouble with the index, so opted to use for loops instead of higher order functions
            //to update sheet link names for renamed sheets
            for (let i = 0; i < renamedSheetIds.length; i++) {
                const sheetData = renamedSheetIds[i];
                for (let j = 0; j < links.length; j++) {
                    const link = links[j];

                    //Get the sheet link url from the richtextvalue
                    const linkUrl = link.getLinkUrl();

                    //Extract the sheet id from the richtextvalue url
                    const sheetIdFromLink = myToc.getSheetGIDFromRichText(linkUrl);
                    //Find the sheet link id that matches the sheet that has been renamed


                    ///////////LOOSE EQUALITY IMPLEMENTED TO COMPARE PARSED DATA////////////////
                    if (sheetIdFromLink == sheetData.id) {
                        sheetData.row = rangeContents.getRow() + j;
                        sheetData.column = rangeContents.getColumn();
                        storedSheetData[sheetData.id].name = sheetData.newName;
                        break;
                    }
                }
                try {
                    //Create new link with updated name
                    newLink = spreadsheetUtil.createSheetLink(sheetData.id, sheetData.url, false, true)
                    //Replace old link with new link in the target cell in the targeted TOC sheet
                    sheet.getRange(sheetData.row, sheetData.column).setRichTextValue(newLink);
                } catch (err) {
                    console.error(`Unable to update link for the sheet renamed: ${sheetData.newName} 
                        with id ${sheetData.id}`
                        , err.stack);
                }
            }
        }
    } catch (err) {
        console.error("Problem setting link", err.stack)
    }

    function getRenamedSheetIds(changeType, myToc, spreadsheetUtil) {
        if (changeType === "OTHER") {
            const currentSheetIds = spreadsheetUtil.getSheets().map(sheet => sheet.getSheetId());
            const renamedSheetIds = [];
            const storedSheetData = myToc.getSheetDataById();

            currentSheetIds.forEach(id => {
                const sheet = spreadsheetUtil.getSheetById(id);
                const sheetName = sheet.getName();
                if (storedSheetData[id]) {
                    const storedSheet = storedSheetData[id];
                    console.log(`STOREDSHEET NAME: ${storedSheet.name} !== CURRENTNAME ${sheetName} ${storedSheet.name !== sheetName}`)
                    if (storedSheet.name !== sheetName) {
                        renamedSheetIds.push({
                            //storedSheet = {id, name, url}
                            ...storedSheet,
                            newName: sheetName
                        });
                    }
                }
            });

            if (renamedSheetIds.length) {
                // console.log("RENAMED SHEETS: ", renamedSheetIds);
                // console.log("myToc stored data by id: ", myToc.sheetDataById);
                return renamedSheetIds;
            }
        }
        return [];
    }
}
