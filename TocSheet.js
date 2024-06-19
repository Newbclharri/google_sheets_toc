class TocSheet {
  
  constructor(name, doSort, titles, sheet, spreadsheet){
    if(TocSheet.instance){
      return (TocSheet.instance)
    }
    this.name = name || "Table of Contents";
    this.spreadsheet = spreadsheet
    this.sheet = sheet;
    this.range.headerName = "TOCHeader";
    this.range.tocName = "TOC";
    this.state.doSort = doSort
    this.titles = titles;
    this.sheetId = sheet.getSheetId();
    this.initialize();
    TocSheet.instance = this;
  }
  static getinstance(){
    if(!TocSheet.instance){
      TocSheet.instance = new TocSheet();
    }
    SpreadsheetUtilIife.active.setActiveSelection(SpreadsheetUtilIife.getSheetById())
    return TocSheet
  }

   setNamedRanges(){
    const numRows = this.titles.length
    this.sheet.setNamedRange(this.range.headerName, sheet.getRange(1,1))
    this.sheet.setNamedRange(this.range.tocName,sheet.getRange(2,1,numRows))
  }

  initialize(){

    //setNamedRange
    this.setNamedRanges()
  }
}