////////////////onEDIT SIMPLE TRIGGER PROVIDES MORE ACCURATE AND DEPENDABLE INFORMATION IN THE EVENT OBJECT//////////////
/////////THAT'S WHY I CHOSE TO BREAK UP EDIT AND CHANGE HANDLING INTO SEPARATE FILES
function onEdit(e) {

    const loaded = TocSheet.load();
    const spreadsheetUtil = SpreadsheetUtility.getInstance();
    const propsStorage = PropertiesServiceStorage.getInstance();
    if (loaded) {
        const myToc = new TocSheet(loaded, spreadsheetUtil, propsStorage)
        const tocRangeHandler = new NamedRangeHandler(e.source.name, e.range, myToc)
        tocRangeHandler.handleRangeEdit(e)
    }
}