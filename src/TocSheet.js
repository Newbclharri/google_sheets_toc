class TocSheet {
  
  constructor({name, titles, sheetId, rangeHeader, rangeToc}, spreadsheet, propsService){
    this.spreadsheet = spreadsheet || new SpreadsheetUtility();
    this.propsService = propsService || new PropertiesServiceStorage();
    this.key = "tocSheet";
    this.name = name|| "Table of Contents";
    this.titles = titles || null;
    this.sheetId = sheetId || null;
    this.rangeHeader = rangeHeader;
    this.rangeToc = rangeToc;
    //this.initialize();
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

  static isEmptyObject(obj){
    if(obj){
      return Object.entries(obj).length === 0 ;
    }
  }

  save(){
    this.propsService.save(this)
  }

  initialize(){
    this.createSheet();
    this.setContentIds();
    this.setTitles();
    this.setSheetLinks();
    this.setNamedRanges();
    this.formatSheet();
  }
  
  formatSheet(){    
    this.pasteHeader();
    this.setFrozenRows(1);
    this.pasteSheetLinks();
  }
  
  
  createSheet(){
    const sheet = this.spreadsheet.insertSheet(this.name)
    this.sheet = sheet;
    this.sheetId = sheet.getSheetId();
  }
  
  setSheetId(){    
    this.sheetId = this.sheet.getSheetId();    
  }

  setContentIds(){
    if(!this.contentIds){
      this.contentIds = this.spreadsheet.getSheetIdsNotEqualTo(this.sheetId);
    }
  }
  setSheetLinks(){
    this.links = this.spreadsheet.createSheetLinks(this.contentIds, false, true)
  }
  
  

  setTitles(titles = null){
    if(titles || this.titles){
      this.titles = titles
    }else{
      const sheets = this.spreadsheet.getSheets();
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
    this.spreadsheet.setNamedRange(this.rangeHeaderName, rangeHeader)
    this.spreadsheet.setNamedRange(this.rangeTocName, rangeToc)
  };

  getSheet(){
    //if sheet object is not empty return this.sheet
    if(!TocSheet.isEmptyObject(this.sheet) && this.sheet){
      return this.sheet;
    }else{
      //if the sheet object is empty getSheet from SpreadsheetUtility
      return this.spreadsheet.getSheetById(this.sheetId);
    }
  }

  getRangeByName(name){
    return this.spreadsheet.getRangeByName(name)
  }

  pasteHeader(){
    const rangeHeader = this.getRangeByName(this.rangeHeaderName);
    rangeHeader
      .setValue(this.name)
      .setFontWeight("bold");
  }

  pasteSheetLinks(){
    const links = this.links;
    const rangeToc = this.getRangeByName(this.rangeTocName);
    rangeToc.setRichTextValues(links);

  }

  setFrozenRows(num){
    if(!TocSheet.isEmptyObject(this.sheet)){
      this.sheet.setFrozenRows(num);
    }else{
      const sheet = this.getSheet();
      sheet.setFrozenRows(num);
    }
  }
  remove(){
    const sheet = this.getSheet();
    this.spreadsheet.deleteSheet(sheet);
    this.propsService.deleteSheetProp(this.key);
  }
  
  handleMenuSelectRemove(){
    this.remove();
  }
}