function onChange(e) {
    let loaded, spreadsheetUtil, propsStorage, backupData, updates, myToc
    console.log("changeType: ", e.changeType)
    if (e.changeType) {
        loaded = TocSheet.load();
        spreadsheetUtil = SpreadsheetUtility.getInstance();
        propsStorage = PropertiesServiceStorage.getInstance();
        let tocChangeHandler;
        if (loaded) {
            myToc = new TocSheet(loaded, spreadsheetUtil, propsStorage)

            switch (e.changeType) {
                case "INSERT_GRID":
                    break;
                case "REMOVE_GRID":
                    const gridHandler = new HandleGridChange(myToc);
                    gridHandler.handleRemoveGrid();
                    break;
                default: //"INSERT_COLUMN, REMOVE_COLUMN,  INSERT_ROW, REMOVE_ROW, OTHER, EDIT"
                    tocChangeHandler = new SheetChangeHandler(myToc)
                    tocChangeHandler.handleChange(e)
                    updates = tocChangeHandler.getSheetUpdates()
                    backupData = tocChangeHandler.getBackupData();
            }

            if(updates && backupData){
                propsStorage.save(myToc.key, updates);
                propsStorage.save(myToc.backupKey, backupData);
            }
        }
    }

}