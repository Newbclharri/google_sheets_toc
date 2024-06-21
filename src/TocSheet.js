class TocSheet {
  
  constructor({name, titles, sheetId, rangeHeader, rangeToc}, spreadsheetApp, propsService){
    //console.log("From Toc: Spreadsheet: ", spreadsheetApp, propsService)
    this.spreadsheetApp = spreadsheetApp || new SpreadsheetUtility();
    this.propsService = propsService || new PropertiesServiceStorage();
    this.state = {};
    this.name = name|| "Table of Contents";
    this.titles = titles || null;
    this.sheetId = sheetId || null;
    this.rangeHeader = rangeHeader;
    this.rangeToc = rangeToc;
    this.initialize();
  }

  static load(){
    const storage = new PropertiesServiceStorage();
    return storage.load();
  }
  static setTocAsActiveSheet(toc){
    const ssApp = new SpreadsheetUtility();
    const sheet = ssApp.getSheetById(toc.sheetId);
    ssApp.setActiveSheet(sheet);
  }

  save(){
    this.propsService.save(this)
  }

  initialize(){
    this.createSheet();
    this.setSheetId();
    this.setTitles();
    this.setNamedRanges();
    //setTitleLinks();
  }
  
  
  createSheet(){
    const sheet = this.spreadsheetApp.insertSheet(this.name)
    this.sheet = sheet;
  }

  setSheetId(){

    if(!this.sheetID){
      this.sheetId = this.sheet.getSheetId();
    }
  }

  setTitles(titles = null){
    if(titles || this.titles){
      this.titles = titles
    }else{
      const sheets = this.spreadsheetApp.getSheets();
      const sheetId = this.sheetId;
      const titles = sheets.filter(sheet => sheet.getSheetId() !== sheetId)
                      .map(sheet => sheet.getName());
      this.titles = titles;
    }
  }

  setNamedRanges(){
    const rangeHeader = this.sheet.getRange(1,1)
    const numRows = this.titles.length;
    const rangeToc = this.sheet.getRange(2,1,numRows)
    if(!this.rangeHeaderName) this.rangeHeaderName = "TOCHeader";
    if(!this.rangeTocName) this.rangeTocName = "TOC";
    this.spreadsheetApp.setNamedRange(this.rangeHeaderName, rangeHeader)
    this.spreadsheetApp.setNamedRange(this.rangeTocName, rangeToc)
  };

}