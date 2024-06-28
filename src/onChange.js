function onChange(e) {
    let loaded, spreadsheetUtil, propsStorage, myToc
    console.log("changeType: ", e.changeType)
    if (e.changeType) {
        loaded = TocSheet.load();
        spreadsheetUtil = SpreadsheetUtility.getInstance();
        propsStorage = PropertiesServiceStorage.getInstance();
        let tocChangeHandler;
        if (loaded) {
            //GET TABLE OF CONTENTS SHEEET ID
            const tocSheetId = propsStorage.load("tocSheetId") || null;
            if(tocSheetId) TocSheet.convertToNumber(tocSheetId);
            myToc = new TocSheet(loaded, spreadsheetUtil, propsStorage);

            switch (e.changeType) {
                case "INSERT_GRID":
                    break;
                case "REMOVE_GRID":
                    const gridHandler = new GridChangeHandler(myToc, tocSheetId);
                    gridHandler.handleRemoveGrid();
                    //updates holder to be saved at the bottom
                    break;
                default: //"INSERT_COLUMN, REMOVE_COLUMN,  INSERT_ROW, REMOVE_ROW, OTHER, EDIT"
                    tocChangeHandler = new SheetChangeHandler(myToc,tocSheetId)
                    tocChangeHandler.handleChange(e)
                    // updates = tocChangeHandler.getSheetUpdates()
                    // backupData = tocChangeHandler.getBackupData();
            }

            // if(updates && backupData){
            //     propsStorage.save(myToc.key, updates);
            //     propsStorage.save(myToc.backupKey, backupData);
            // }
        }
    }

}