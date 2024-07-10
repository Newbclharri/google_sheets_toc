function onOpen() {
  const uI = UiUtil.getInstance();
  uI.createMenu();
  
  const loaded = TocSheet.load();

  if(loaded){

    const myToc = new TocSheet(loaded);

    console.log("CLEAN UP ON OPEN: ", myToc.cleanUpSheetDataById());

    myToc.save();
    myToc.saveBackup();
  }
 
}
