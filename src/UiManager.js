class UiUtil{
  constructor(spreadsheet){
    this.uI = spreadsheet().getUi();
  }

  alert(message) {
    return this.uI.alert(message);
  }

  createMenu(){
    return this.uI.createAddonMenu()
      .addItem("Insert", "start")
      .addSeparator()
      .addItem("Remove","confirmDelete")
      .addToUi();
    }

  confirmDelete(){
    const response = this.uI.alert('Confirm Deletion', 'Delete this item?', this.ui.ButtonSet.OK_CANCEL);

    if (response) {
      return true;
    } else{
      this.uI.alert('Deletion Cancelled');
    }
    return false;
  }
}

function start(){
  const uI = getSpreadsheetApp().getUi()
  const loaded = TocSheet.load();
  if(loaded){
    TocSheet.setTocAsActiveSheet(loaded);
  }else{
    const spreadsheetApp = new SpreadsheetUtility();
    const propsService = new PropertiesServiceStorage();
    const myToc = new TocSheet({},spreadsheetApp, propsService);
    myToc.save();
    uI.alert("sheet created");    
  }  
}

// function setFalse(){
//   const doSort = false
//   const tocSheet = new SheetManager("TOC", getSpreadsheetApp, getPropsService,getScriptApp);
// }

function confirmDelete(){
  const isConfirmed = Ui.confirmDelete;
  if(isConfirmed){
    removeToc(isConfirmed)
  }
}