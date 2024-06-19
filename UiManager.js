class UiUtil{
    constructor(spreadsheet){
      this.uI = spreadsheet().getUi();
    }

    alert(message) {
      return this.uI.alert(message);
    }

    createMenu(){
      return this.uI.createAddonMenu()
        .addItem("Sorted", "setTrue")
        .addItem("Unsorted","setFalse")
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

  function setTrue(){
    const uI = getSpreadsheetApp().getUi()
    const doSort = true
    const sheet = JSON.parse(getPropsService().getScriptProperties().getProperty("sheet"))
    console.log("sheet from props service: ", sheet)
    if(!sheet){
      const tocSheet = new SheetManager("TOC", getSpreadsheetApp, getPropsService,getScriptApp);
      uI.alert("sheet created")
      console.log(tocSheet)
    }
    
  }

  function setFalse(){
    const doSort = false
    const tocSheet = new SheetManager("TOC", getSpreadsheetApp, getPropsService,getScriptApp);
  }

  function confirmDelete(){
    const isConfirmed = Ui.confirmDelete;
    if(isConfirmed){
      removeToc(isConfirmed)
    }
  }