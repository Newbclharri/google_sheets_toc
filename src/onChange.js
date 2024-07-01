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
                    break;
                case "REMOVE_GRID":
                    handleGridChange(myToc, sheetId, "REMOVE_GRID");
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