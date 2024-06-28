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

class GridChangeHandler {
    constructor(obj = {}, id) {
        this.tab = obj;
        this.tocSheetId = id
        this.ssUtil = SpreadsheetUtility.getInstance();
        this.uI = UiUtil.getInstance(); //SpreadsheetApp.getUi();
        this.currentListOfSheetIds = this.tab.fetchSheetIds();
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
        console.log("allIds: ", "from property: this.tab.sheetId: ", this.tab.sheetId, " from gridhandler.currentIdLIst: ", this.currentListOfSheetIds)
        console.log(this.tab.fetchSheetIds().some(id => id === this.tocSheetId))
        return !this.tab.fetchSheetIds().some(id => id === this.tocSheetId)
    }

    handleUserRemovesTocTab() {
        // const key = this.tab.backupKey;
        const backup = this.tab.toJSON();
        //logic
        //update to potentially re-add links
        this.tab.updateState({"allSheetIds": this.currentListOfSheetIds})
        this.tab.restore(backup);
        this.uI.alert("Select 'Remove' from Table of Contents menu to remove this sheet.")

    }

    handleUserRemovesContentTab() {
        //logic
    }

    handleUserInsertsContentTab(){
        //logic
    }

}