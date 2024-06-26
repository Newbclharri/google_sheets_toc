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
    // this.rangeHeaderName = rangeHeaderName || null;
    // this.rangeContentsName = rangeContentsName || null;
    // this.rangeHeaderA1Notation = rangeHeaderA1Notation || null;
    // this.rangeContentsA1Notation = rangeContentsA1Notation || null;
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

  save(key, data) {
    this.propsStorage.save(key, data);
  }

  initialize() {
    this.createSheet();
    this.initializeAllSheetIds();
    this.initializeContentIds();
    this.initializeTitles();
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
    const sheet = this.spreadsheetUtil.insertSheet(this.name);
    this.sheet = sheet;
    this.sheetId = sheet.getSheetId();
  }

  initializeSheetId() {
    this.sheetId = this.sheet.getSheetId();
  }

  setContentIds(sheetIds) {
    this.contentIds = sheetIds;
  }

  initializeAllSheetIds() {
    //all sheet Ids in the active spreadsheet for reference
    this.allSheetIds = this.spreadsheetUtil.getSheetIds();

  }

  initializeContentIds() {
    //this.contentIds = this.spreadsheetUtil.getSheetIdsNotEqualTo(this.sheetId);
    this.contentIds = this.allSheetIds.filter(id => id !== this.sheetId);
  }

  createSheetLinks(contentIds = this.contentIds) {
    this.links = this.spreadsheetUtil.createSheetLinks(this.contentIds, false, true);
  }



  initializeTitles() {
    try {
      if (!this.titles) {
        const sheets = this.spreadsheetUtil.getSheets();
        const sheetId = this.sheetId;
        const titles = sheets.filter(sheet => sheet.getSheetId() !== sheetId)
          .map(sheet => sheet.getName());
        this.titles = titles;
      }

    } catch (err) {
      console.log(err)
      this.titles = [];
    }

  }

  initializeNamedRanges() {
    const rangeHeader = this.sheet.getRange(1, 1)
    const numRows = this.titles.length;
    const rangeContents = this.sheet.getRange(2, 1, numRows)
    if (!this.rangeHeaderName) this.rangeHeaderName = "TOCHeader";
    if (!this.rangeContentsName) this.rangeContentsName = "TOC";
    this.spreadsheetUtil.setNamedRange(this.rangeHeaderName, rangeHeader)
    this.spreadsheetUtil.setNamedRange(this.rangeContentsName, rangeContents)
    this.rangeHeader = rangeHeader;
    this.rangeContents = rangeContents;
    this.rangeHeaderA1Notation = rangeHeader.getA1Notation();
    this.rangeContentsA1Notation = rangeContents.getA1Notation();
  };


  loadData() {
    const sheet = this.spreadsheetUtil.getSheetById(this.sheetId);
    const data = sheet.getDataRange().getValues();
    this.data = data;
  }

  getRichTextValues() {
    const rangeContents = this.spreadsheetUtil.getRangeByName(this.rangeContentsName);
    this.richTextValues = rangeContents.getRichTextValues();
    console.log("richTextValues: ", this.richTextValues)
  }

  loadBackup() {
    const headerRangeA1Notation = this.spreadsheetUtil.getActive()
      .getRangeByName(this.rangeHeaderName)
      .getA1Notation();
    const contentsRangeA1Notation = this.spreadsheetUtil.getActive()
      .getRangeByName(this.rangeContentsName)
      .getA1Notation();
    let key = "tocBackup", name = this.name, data, richTextValues;
    this.loadData();
    //this.setRichTextValues();
    data = this.data;
    // richTextValues = this.richTextValues;
    this.backup = {
      key,
      name,
      headerRangeA1Notation,
      contentsRangeA1Notation,
      data,
    }
    this.save(this.backup.key, this.backup)
  }

  getBackUp() {
    return this.backup
  }

  getSheet() {
    //if sheet object is not empty return this.sheet
    if (this.sheet && !TocSheet.isEmptyObject(this.sheet)) {
      return this.sheet;
    } else {
      //if the sheet object is empty, undefined, or null getSheet from SpreadsheetUtility
      const sheet = this.spreadsheetUtil.getSheetById(this.sheetId);
      return sheet;
    }
  }

  getSheetId() {
    if (this.sheetId) {
      return this.sheetId
    }
    return undefined
  }

  getAllSheetIds() {
    try {

      if (this.allSheetIds === null) {
        return this.spreadsheetUtil.getAllSheetIds()
      }
      return this.allSheetIds;
    } catch (err) {
      console.log(err)
      return this.allSheetIds = []
    }
  }


  getRangeByName(name) {

    const range = this.spreadsheetUtil.getRangeByName(name)
    if (range) {
      return this.spreadsheetUtil.getRangeByName(name)
    }
    return undefined;
  }

  pasteHeader(rangeHeader) {
    rangeHeader = rangeHeader || this.getRangeByName(this.rangeHeaderName);
    rangeHeader
      .setValue(this.name)
      .setFontWeight("bold");
  }

  pasteSheetLinks() {
    const links = this.links;
    const rangeContents = this.getRangeByName(this.rangeContentsName);
    rangeContents.setRichTextValues(links);

  }

  setFrozenRows(num) {
    if (this.sheet && !TocSheet.isEmptyObject(this.sheet)) {
      this.sheet.setFrozenRows(num);
    } else {
      const sheet = this.getSheet();
      sheet.setFrozenRows(num);
    }
  }
  remove() {
    const sheet = this.getSheet();
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

  restore(backup) {
    console.log(backup)
    if (backup) {
      const newSheet = this.spreadsheetUtil.createSheet(backup.name);
      let range;
      newSheet.getRange(1, 1, backup.data.length, backup.data[0].length).setValues(backup.data)
      range = newSheet.getRange(backup.rangeHeaderA1Notation)
      this.spreadsheetUtil.setNamedRange(this.rangeContentsName, range);
      range = newSheet.getRange(backup.rangeContentsA1Notation)
      this.spreadsheetUtil.setNamedRange(this.rangeContentsName, range)
      this.name = backup.name;
      //this.sheetId = newSheet.getSheetId();
      console.log("newSheetId: ", this.sheetId)

      //destructure if more efficient before saving
      //const {this object property names here} = this;
      this.save("tocSheet", this);
    }
  }
}