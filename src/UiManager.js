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

  confirmDelete(title, message){
    const response = this.uI.alert(title, message, this.uI.ButtonSet.OK_CANCEL);
    if (response == this.uI.Button.OK) {
      return true;
    }else{
      this.uI.alert('Deletion Cancelled');
    }  
      
    
    return false;
  }
}

function start(){
  const scriptApp = getScriptApp(), spreadsheetUtil = SpreadsheetUtility.getInstance();
  const uI = spreadsheetUtil.getUi();
  const loaded = TocSheet.load();
  if(loaded){
    TocSheet.setTocAsActiveSheet(loaded);
  }else{
    const propsStor = new PropertiesServiceStorage();
    const triggerManager = TriggerManager.getInstance(scriptApp, spreadsheetUtil);
    const myToc = new TocSheet({},spreadsheetUtil, propsStor);
    let propsToSave;
    myToc.initialize();
    console.log("From ui manager tocKey: ", myToc.key)
    propsToSave = [["tocSheetId", myToc.sheetId], [myToc.key, myToc.toJSON()], [myToc.backupKey, myToc.getBackUp()]];
    //////////INITIAL SAVE//////////////
    for(let prop of propsToSave){
      const key = prop[0], value = prop[1];
      PropertiesServiceStorage.getInstance().save(key, value);
    }
    triggerManager.setTrigger(triggerManager.getEventType().ON_CHANGE, "onChange");
  }  
}

function confirmDelete(){
  const uI = UiUtil.getInstance()
  const isConfirmed = uI.confirmDelete('Confirm Delete:','Delete item?');
  if(isConfirmed){
    const loaded = TocSheet.load();
    if(loaded){
      const scriptApp = getScriptApp(), spreadsheetUtil = SpreadsheetUtility.getInstance();
      const triggerManager = TriggerManager.getInstance(scriptApp, spreadsheetUtil);
      const myToc = new TocSheet(loaded);
      myToc.handleMenuSelectRemove(()=>{
        return triggerManager.deleteTrigger(triggerManager.getEventType().ON_CHANGE, "handleOnChange");
      });
      uI.alert("deleted");
    }else{
      uI.alert("No Table of Contents found.")
    }
  }
}