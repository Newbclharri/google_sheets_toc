function onChange(e) {
    console.log("changeType: ",e.changeType)
    if (e.changeType !== "EDIT") {
        const loaded = TocSheet.load();
        const spreadsheetUtil = SpreadsheetUtility.getInstance();
        const propsStorage = PropertiesServiceStorage.getInstance();
        let backupData;
        let tocChangeHandler;
        if (loaded) {
            const myToc = new TocSheet(loaded, spreadsheetUtil, propsStorage)

            switch (e.changeType) {
                case "INSERT_GRID":
                    break;
                case "REMOVE_GRID":
                    const gridHandler = new HandleGridChange(myToc);
                    break;
                default: //"INSERT_COLUMN, REMOVE_COLUMN,  INSERT_ROW, REMOVE_ROW, OTHER"
                    tocChangeHandler = new SheetChangeHandler(myToc)
                    tocChangeHandler.handleChange(e)
                    propsStorage.save(myToc.key, tocChangeHandler.getSheetUpdates());
                    backupData = tocChangeHandler.getBackupData();
                    propsStorage.save(myToc.backupKey, backupData);
            }
        }
    }
   
}