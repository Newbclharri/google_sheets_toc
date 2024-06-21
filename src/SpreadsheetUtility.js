class SpreadsheetUtility {
    constructor(){
      this.activeSheet = SpreadsheetApp.getActive();
    }
      insertSheet(name){
        return this.activeSheet.insertSheet(name,0)
      };
  
      setNamedRange(name,range){
        this.activeSheet.setNamedRange(name, range)
      };
  
      getSheets(){
        return this.activeSheet.getSheets()
      };
  
      getSheetByName(name){
        return this.activeSheet.getSheetByName(name)
      };
    }
  
  function myFunction(){
    console.log(new SpreadsheetUtility().getSheets());
}  