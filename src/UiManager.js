class UiUtil{
  constructor(){
    if(UiUtil.instance){
      return UiUtil.instance
    }
    this.uI = getSpreadsheetApp().getUi() || SpreadsheetApp.getUi();
    UiUtil.instance = this;
  }

  static getInstance(){
    if(!UiUtil.instance){
      return new UiUtil();
    }
    return UiUtil.instance;
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
    const response = this.uI.alert('Confirm Deletion', 'Delete this item?', this.uI.ButtonSet.OK_CANCEL);
    if (response == this.uI.Button.OK) {
      return true;
    }else{
      this.uI.alert('Deletion Cancelled');
    }  
      
    
    return false;
  }
}

function start(){
  const uI = getSpreadsheetApp().getUi();
  const loaded = TocSheet.load();
  if(loaded){
    TocSheet.setTocAsActiveSheet(loaded);
  }else{
    const spreadsheetApp = new SpreadsheetUtility();
    const propsService = new PropertiesServiceStorage();
    const myToc = new TocSheet({},spreadsheetApp, propsService);
    myToc.initialize();
    myToc.save();
  }  
}

function confirmDelete(){
  const uI = UiUtil.getInstance()
  const isConfirmed = uI.confirmDelete();
  if(isConfirmed){
    const loaded = TocSheet.load();
    if(loaded){
      const myToc = new TocSheet(loaded);
      myToc.handleMenuSelectRemove();
      uI.alert("deleted");
    }else{
      uI.alert("No Table of Contents found.")
    }
  }
}