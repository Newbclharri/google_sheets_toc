class SpreadsheetUtility {
  constructor(){
    if(SpreadsheetUtility.instance){
      return SpreadsheetUtility.instance
    }
    this.spreadsheetApp = getSpreadsheetApp(); //|| SpreadsheetApp;
    this.activeSheet = this.spreadsheetApp.getActive();
    this.url = this.activeSheet.getUrl();
    this.sheets = this.activeSheet.getSheets();
    this.newRichTextStyle = this.spreadsheetApp.newTextStyle();
    this.newRichTextValue = this.spreadsheetApp.newRichTextValue();
    SpreadsheetUtility.instance = this;
  }

    static getInstance(){
      if(!SpreadsheetUtility.instance){
        return new SpreadsheetUtility();
      }
      return SpreadsheetUtility.instance;
    }

    getActive(){
      return this.activeSheet;
    }

    getUi(){
      this.spreadsheetApp.getUi();
    }

    insertSheet(name){
      return this.activeSheet.insertSheet(name,0);

    }

    setActiveSheet(sheet){
      return this.activeSheet.setActiveSheet(sheet);
    }

    setNamedRange(name,range){
      this.activeSheet.setNamedRange(name, range)
    }

    getSheets(){
      return this.sheets;
    }

    getSheetByName(name){
      return this.activeSheet.getSheetByName(name)
    }
    getSheetById(id){
      if(!isNaN(id)){
        return this.sheets.find(sheet => sheet.getSheetId() === id);
      }
      return undefined;
    }

    getSheetIds(){
      return this.activeSheet.getSheets().map(sheet => sheet.getSheetId());
    }

    getSheetIdsNotEqualTo(tocId){
      const ids = this.sheets.filter(sheet => sheet.getSheetId() !== tocId)
          .map(sheet => sheet.getSheetId());
      return ids

    }

    getRangeByName(name){
      return this.activeSheet.getRangeByName(name)
    }

    getA1Notation(range){
      return range.getA1Notation()
    }

    createSheetLink(sheet, underline=false, bold=false){
      if(sheet){ 
        let sheetId = sheet.getSheetId()
        let sheetUrl = this.url + "?gid=" + sheetId + "#gid=" + sheetId;
        let linkStyle = this.newRichTextStyle
          .setUnderline(underline)
          .setBold(bold)
          .build();
        let link = this.newRichTextValue
          .setText(sheet.getName())
          .setLinkUrl(sheetUrl)
          .setTextStyle(linkStyle)
          .build()
        return link;
      }
    }

    createSheetLinks(sheetIds, underline = false, bold = false){
      //get all sheets from passed sheetIds
      const sheets = [];
      const links = [];
      let sheetUrl, linkStyle, link;
      sheetIds.forEach(id =>{
        const sheet = this.getSheetById(id);
        if(sheet){
          link = this.createSheetLink(sheet, underline, bold);
          links.push([link]);
        }
      }) 
      return links;           
    }

    deleteSheet(sheet){
      this.activeSheet.deleteSheet(sheet);
    }
}