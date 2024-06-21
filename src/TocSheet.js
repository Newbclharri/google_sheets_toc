class TocSheet {
  
  constructor(name, titles, spreadsheetApp, propsService){
    console.log("From Toc: Spreadsheet: ", spreadsheetApp, propsService)
    this.state = {};   
    this.state.name = name || "Table of Contents";
    this.spreadsheetApp = spreadsheetApp || new SpreadsheetUtility();
    this.propsService = propsService || new PropertiesServiceStorage();
    this.state.properties = {
      sheetId: null,
      titles: null,
      range:{
        headerName: "TOCHeader",
        tocName: "TOC"
    }}
    this.initialize();
  }

  static load(){
    const storage = new PropertiesServiceStorage();
    storage.load();
  }

  initialize(){
    this.createSheet();
    this.setSheetId();
    this.setTitles();
    this.setNamedRanges();
    //setTitleLinks();
  }


  createSheet(){

    console.log("name",this.state.name)
    const sheet = this.spreadsheetApp.insertSheet("Table of Contents")
    this.sheet = sheet;
  }

  setSheetId(){
    this.state.properties.sheetId = this.sheet.getSheetId();
  }

  setTitles(titles = null){
    if(titles){
      this.state.properties.titles = titles
    }else{
      const sheets = this.spreadsheetApp.getSheets();
      const titles = sheets.filter(title => title.getSheetId() !== this.sheetId);
      this.state.properties.titles = titles;

    }
  }

  setNamedRanges(){
    const rangeHeader = this.sheet.getRange(1,1)
    const numRows = this.state.properties.titles.length;
    const rangeToc = this.sheet.getRange(2,1,numRows)
    this.spreadsheetApp.setNamedRange(this.state.properties.range.headerName, rangeHeader)
    this.spreadsheetApp.setNamedRange(this.state.properties.range.tocName, rangeToc)
  }

}