class SpreadsheetUtility {
  constructor(){
    this.spreadsheetApp = SpreadsheetApp;
    this.activeSheet = this.spreadsheetApp.getActive();
    this.url = this.activeSheet.getUrl();
    this.sheets = this.activeSheet.getSheets();
    this.newRichTextStyle = this.spreadsheetApp.newTextStyle();
    this.newRichTextValue = this.spreadsheetApp.newRichTextValue();
  }
    insertSheet(name){
      return this.activeSheet.insertSheet(name,0)
    };

    setActiveSheet(sheet){
      return this.activeSheet.setActiveSheet(sheet);
    };

    setNamedRange(name,range){
      this.activeSheet.setNamedRange(name, range)
    };

    getSheets(){
      return this.sheets;
    };

    getSheetByName(name){
      return this.activeSheet.getSheetByName(name)
    };
    getSheetById(id){
      if(!isNaN(id)){
        const sheet = this.sheets.find(sheet => sheet.getSheetId() === id);
        return sheet;
      }
      return undefined;
    };

    getSheetIdsNotEqualTo(tocId){
      const ids = this.sheets.filter(sheet => sheet.getSheetId() !== tocId)
          .map(sheet => sheet.getSheetId());
      return ids

    }

    getRangeByName(name){
      return this.activeSheet.getRangeByName(name)
    }

    createSheetLink(sheet, underline, bold){
      if(sheet){ 
        let sheetUrl = this.url + "#gid" + sheet.getSheetId();
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
          link = this.createSheetLink(sheet, true, true);
          links.push([link]);
        }
      }) 
      return links;           
    }
}