////////////////onEDIT SIMPLE TRIGGER PROVIDES MORE ACCURATE AND DEPENDABLE INFORMATION IN THE EVENT OBJECT//////////////
/////////THAT'S WHY I CHOSE TO BREAK UP EDIT AND CHANGE HANDLING INTO SEPARATE FILES
function onEdit(e, changeType) {
    console.log("CHANGE TYPE: ", changeType);
    const spreadsheetUtil = SpreadsheetUtility.getInstance();
    const propsStorage = new PropertiesServiceStorage();
    const sheetId = propsStorage.load("tocSheetId");
    let myToc;
    let tocSheetDoesExist = false;

    if (sheetId) {
        try {
            const loaded = TocSheet.load();
            myToc = new TocSheet(loaded, spreadsheetUtil, propsStorage);
            tocSheetDoesExist = myToc.doesExistSheet();
        } catch (err) {
            console.log("An error occured attempting to find the TOC sheet: ", err);
        }
    }

    if (tocSheetDoesExist) {
        console.log("SHEET DOES EXIST. WE CAN DO WORK!")
        try{
            handleRangeEdit(e, myToc)
        }catch(err){
            console.error(err)
            console.log(err.stack)
        }
    } else {
        console.log("SHEET DOES NOT EXIST. CAN'T DO WORK!");
    }
}

function handleRangeEdit(e, myToc) {
    const tocRangeHandler = new NamedRangeHandler(e.source.name, e.range, myToc)
    tocRangeHandler.handleRangeEdit(e)

}