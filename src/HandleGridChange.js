// class OnChangeHandler{
//     constructor(spreadsheetUtil){
//         this.REMOVE_GRID = "REMOVE_GRID";
//         //this.tocSheet = tocSheet;
//         this.spreadsheetUtil = spreadsheetUtil // || SpreadsheetApp
//     }



//     handleRemoveGrid(changeType, callback = null){
//         if(changeType === this.REMOVE_GRID){

//             if(callback){
//                 callback();
//             }
//         }
//     }
// }

class HandleGridChange {
    constructor(obj) {
        this.instance = obj;
        this.ssUtil = SpreadsheetUtility.getInstance();
        this.uI = UiUtil.getInstance(); //SpreadsheetApp.getUi();
        this.currentListOfSheetIds = this.instance.fetchSheetIds();
        this.propsStorage = PropertiesServiceStorage.getInstance(); //PropertiesService
        //this.handleRemoveGrid();

    }

    handleRemoveGrid() {

        if (this.isRemovedTocTab()) {
            this.handleUserRemovesTocTab();
        }
    }

    isRemovedTocTab() {
        //logic here
        console.log("allIds: ", "from property: this.instance.sheetId: ", this.instance.sheetId, " from gridhandler.currentIdLIst: ", this.currentListOfSheetIds)
        console.log(this.instance.fetchSheetIds().some(id => id === this.instance.sheetId))
        return !this.instance.fetchSheetIds().some(id => id === this.instance.sheetId)
    }

    handleUserRemovesTocTab() {
        const key = this.instance.backupKey;
        const backup = this.propsStorage.load(key);
        console.log(backup);
        //logic
        this.instance.updateState({"allSheetIds": this.currentListOfSheetIds})
        this.instance.restore(backup);
        this.uI.alert("Table of Contents Removed.")

    }

    handleUserRemovesContentTab() {
        //logic
    }

}