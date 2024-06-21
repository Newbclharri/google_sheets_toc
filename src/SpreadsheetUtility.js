class SpreadsheetUtility {
    constructor(){
      this.activeSheet = SpreadsheetApp.getActive();
      this.sheets = this.activeSheet.getSheets();
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
        if(id){
          const sheet = this.sheets.find(sheet => sheet.getSheetId() === id);
          return sheet;
        }
        return undefined;
      };
  }