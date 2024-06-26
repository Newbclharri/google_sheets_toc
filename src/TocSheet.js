class TocSheet {

  constructor({ name, titles, sheetId, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation }, spreadsheetUtil, propsStorage) {
    this.sheetId = sheetId || null;
    this.allSheetIds = allSheetIds || null;
    this.spreadsheetUtil = spreadsheetUtil || SpreadsheetUtility.getInstance();
    this.propsStorage = propsStorage || PropertiesServiceStorage.getInstance();
    this.key = "tocSheet";
    this.backupKey = "tocBackup";
    this.name = name || "Table of Contents";
    this.titles = titles || null;
    this.rangeHeaderName = rangeHeaderName || null;
    this.rangeContentsName = rangeContentsName || null;
    this.rangeHeaderA1Notation = rangeHeaderA1Notation || null;
    this.rangeContentsA1Notation = rangeContentsA1Notation || null;
  }

  static load() {
    const storage = PropertiesServiceStorage.getInstance();
    return storage.load();
  }
  static setTocAsActiveSheet(toc) {
    const ssUtil = SpreadsheetUtility.getInstance();
    const sheet = ssUtil.getSheetById(toc.sheetId);
    ssUtil.setActiveSheet(sheet);
  }

  static isEmptyObject(obj) {
    if (obj && typeof obj === "object") {
      if (!Array.isArray(obj) && Object.entries(obj).length === 0) {
        return true;
      }
    }
    return false;
  }

  save(key = this.key, data = this) {
    this.propsStorage.save(key, data);
  }

  initialize(name = this.name) {
    this.createSheet(name);
    this.allSheetIds = this.fetchSheetIds();
    this.contentIds = this.fetchSheetIdsNotEqualTo(this.sheetId)
    this.titles = this.fetchTitleNames();
    this.createSheetLinks();
    this.initializeNamedRanges();
    this.formatSheet();
  }

  formatSheet() {
    this.pasteHeader();
    this.setFrozenRows(1);
    this.pasteSheetLinks();
  }


  createSheet(name) {
    name = name || this.name;
    const sheet = this.spreadsheetUtil.insertSheet(this.name)
    this.sheet = sheet;
    this.sheetId = sheet.getSheetId();
    return sheet

  }

  initializeSheetId() {
    this.sheetId = this.sheet.getSheetId();
  }

  setContentIds(sheetIds) {
    this.contentIds = sheetIds;
  }

  setAllSheetIds(sheetIds) {
    this.allSheetIds = sheetIds;
  }

  fetchSheetIds() {
    return this.spreadsheetUtil.getSheetIds();
  }

  updateSheetIds() {
    this.allSheetIds = this.fetchSheetIds();
  }

  initializeAllSheetIds() {
    //all sheet Ids in the active spreadsheet for reference
    this.allSheetIds = this.spreadsheetUtil.getSheetIds();

  }

  initializeContentIds() {
    //this.contentIds = this.spreadsheetUtil.getSheetIdsNotEqualTo(this.sheetId);
    this.contentIds = this.fetchSheetIdsNotEqualTo(this.sheetId)
  }

  fetchSheetIdsNotEqualTo(id = this.sheetId) {
    return this.spreadsheetUtil.getSheetIdsNotEqualTo(id);
  }

  createSheetLinks(contentIds = this.contentIds) {
    if (!(contentIds && contentIds.length)) {
      contentIds = this.fetchSheetIdsNotEqualTo(this.sheetId)
    }
    this.links = this.spreadsheetUtil.createSheetLinks(contentIds, false, true);
    return this.links;
  }



  fetchTitleNames() {
    let titles;
    try {
      const sheets = this.spreadsheetUtil.getSheets();
      const sheetId = this.sheetId;
      titles = sheets.filter(sheet => sheet.getSheetId() !== sheetId)
        .map(sheet => sheet.getName());
    } catch (err) {
      console.log(err)
      titles = [];
    }
    return titles;

  }

  initializeNamedRanges() {

    const rangeHeader = this.rangeHeaderA1Notation ? this.sheet.getRange(this.rangeHeaderA1Notation) : this.sheet.getRange(1, 1)

    const numRows = this.titles.length;
    const rangeContents = this.sheet.getRange(2, 1, numRows)
    if (!this.rangeHeaderName) this.rangeHeaderName = "TOCHeader";
    if (!this.rangeContentsName) this.rangeContentsName = "TOC";
    this.setNamedRange(this.rangeHeaderName, rangeHeader)
    this.setNamedRange(this.rangeContentsName, rangeContents)
    this.rangeHeader = rangeHeader;
    this.rangeContents = rangeContents;
    this.rangeHeaderA1Notation = rangeHeader.getA1Notation();
    this.rangeContentsA1Notation = rangeContents.getA1Notation();
  };

  setNamedRange(range, name) {
    this.spreadsheetUtil.setNamedRange(range, name)
  }


  loadData() {
    const sheet = this.spreadsheetUtil.getSheetById(this.sheetId);
    const data = sheet.getDataRange().getValues();
    this.data = data;
  }

  getRichTextValues() {
    const rangeContents = this.spreadsheetUtil.getRangeByName(this.rangeContentsName);
    this.richTextValues = rangeContents.getRichTextValues();
    //console.log("richTextValues: ", this.richTextValues)
  }

  saveBackup() {
    const { name, backupKey, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation } = this;
    this.backup = { name, backupKey, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation };
    this.save(this.backupKey, this.backup);
  }

  getBackUp() {
    return this.backup
  }

  getSheet() {
    try {
      //if sheet object is not empty return this.sheet
      if (this.sheet && !TocSheet.isEmptyObject(this.sheet)) {
        return this.sheet;
      }

    } catch (err) {
      console.log(err)
      return undefined;
    }
  }

  fetchSheet() {
    return this.spreadsheetUtil.getSheetById(this.sheetId)
  }

  getSheetId() {
    if (this.sheetId) {
      return this.sheetId
    }
    return undefined
  }

  getAllSheetIds() {
    try {
      return this.allSheetIds;
    } catch (err) {
      console.log(err)
      return this.allSheetIds = []
    }
  }



  getRangeByName(name) {
    const range = this.spreadsheetUtil.getRangeByName(name)

    if (range) {
      return range
    }
    return undefined;
  }

  pasteHeader(range) {
    range = range || this.getRangeByName(this.rangeHeaderName);
    let sheet;
    if (!range) {
      try {
        sheet = this.sheet;
      } catch (error) {
        sheet = this.fetchSheet();
        console.log(error.message, " | ", error.stack)
      }finally{
        if(sheet){
          range = sheet.getRange(this.rangeHeaderA1Notation)
          this.setNamedRange(this.rangeHeaderName, range)
          range = this.getRangeByName(this.rangeHeaderName)
          range
            .setValue(this.name)
            .setFontWeight("bold");
        }
      }
    }else{
      console.log("Could not retrieve sheet.")
    }
  }

  pasteSheetLinks(range, links = this.links) {
    
    range = range || this.getRangeByName(this.rangeContentsName);

    //if links is not an arrat with elements
    if (!(links && links.length)) {
      links = this.createSheetLinks()
    }

    if (!range) {
      let sheet;
      try {
        sheet = this.sheet;
      } catch (err) {
        sheet = this.fetchSheet();
        console.error(err.message, " | ", err.stack)
      } finally {
        if(sheet){
          range = sheet.getRange(this.rangeContentsA1Notation)
          this.setNamedRange(this.rangeContentsName, range)
          range = this.getRangeByName(this.rangeContentsName)
          range.setRichTextValues(links);
        }
      }
    }else{
      console.error("Couldn't retrieve sheet.")
    }

  }

  setFrozenRows(num) {
    let sheet = this.sheet;
    //if the sheet is an empty object or undefined fetch the TOC sheet from the spreadsheet
    try {
      if ((sheet && TocSheet.isEmptyObject(sheet)) || !sheet) {
        sheet = this.fetchSheet();
      }
      sheet.setFrozenRows(num)
    } catch (err) {
      console.log(err, `| Could not find sheet with id . ${this.sheetId}`)
    }
  }
  remove() {
    const sheet = this.fetchSheet();
    if (sheet) {
      this.spreadsheetUtil.deleteSheet(sheet);
    }
    this.propsStorage.deleteSheetProp(this.key);
  }

  handleMenuSelectRemove(callback = null) {
    this.remove();
    callback();
  }

  handleUserDeletesTocTab(toc, callback = null) {
    // const targetId = this.sheetId === undefined ? toc.sheetId : this.sheetId;
    // const allSheetIds = this.spreadsheetUtil.getSheetIds();
    // if(!allSheetIds.some(id => id === targetId)){
    //   this.remove();
    // }
    if (callback) {
      callback();
    }
  }

  updateState(newState = {}) {
    for (let key in newState) {
      if (this.hasOwnProperty(key)) {
        this[key] = newState[key]
      }
    }
  }

  restore(backup) {
    //console.log(backup)
    if (backup && !TocSheet.isEmptyObject(backup)) {
      const { rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation } = backup
      let newSheet;
      this.updateState({ rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation })
      this.name = backup.name;
      if (backup.data) {
        newSheet = this.createSheet(backup.name);
        newSheet.getRange(1, 1, backup.data.length, backup.data[0].length).setValues(backup.data)
        this.sheetId = newSheet.getSheetId();
      } else {
        this.createSheet(backup.name);
      }
      this.formatSheet();
      console.log("newSheetId: ", this.sheetId)
      //const {name, titles, sheetId, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation} = this;
      //dataToSave = {name, titles, sheetId, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation}
      //destructure if more efficient before saving
      //const {this object property names here} = this;
      //this.save("tocSheet", this);
    }
  }
}