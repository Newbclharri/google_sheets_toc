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
        this.uI = UiUtil.getInstance(); //SpreadsheetApp.getUi();
        this.handleRemoveGrid();

    }

    handleRemoveGrid() {

        if (this.isRemovedTocTab()) {
            this.handleUserRemovesTocTab()
        }
    }

    isRemovedTocTab() {
        //logic here
        return this.instance.getAllSheetIds().some(id => id === this.instance.sheetId)
    }

    handleUserRemovesTocTab() {
        //logic
        this.uI.alert("Table of Contents Removed.")
    }

    handleUserRemovesContentTab() {
        //logic
    }

}