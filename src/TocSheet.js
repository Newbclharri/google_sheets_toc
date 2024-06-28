class TocSheet {

  constructor(params = {}, spreadsheetUtil, propsStorage) {
    this.sheetId = params.sheetId || null;
    this.allSheetIds = params.allSheetIds || null;
    this.spreadsheetUtil = spreadsheetUtil || SpreadsheetUtility.getInstance();
    this.propsStorage = propsStorage || PropertiesServiceStorage.getInstance();
    this.key = "tocSheet";
    this.backupKey = "tocBackup";
    this.tocSheetIdKey = "tocSheetId";
    this.name = params.name || "Table of Contents";
    this.titles = params.titles || null;
    this.rangeHeaderName = params.rangeHeaderName || null;
    this.rangeContentsName = params.rangeContentsName || null;
    this.rangeHeaderA1Notation = params.rangeHeaderA1Notation || null;
    this.rangeContentsA1Notation = params.rangeContentsA1Notation || null;
    this.dataRangeValues = params.dataRangeValues;
  }

  static isValidJson(str) {
    try {
      JSON.parse(str);
      return true;
    } catch (e) {
      return false;
    }
  }

  static convertToJson(data) {
    if (this.isValidJson(data)) {
      return data;
    }
    return JSON.stringify(data);
  }

  static convertToNumber(str) {
    //Return string if it is actually a number
    if(!isNaN(str)){
      return str
    }

    if (typeof str !== "string") {
      throw new Error(`Value ${str} is not a string. It is Type ${typeof str}`);
    }

    const num = Number(str)
    if (Number.isNaN(num)) {
      throw new Error(`Value ${str} is not coercible to a number.`);
    }

    return num;
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

  save(key = this.key, data = null) {
    data = data || this.toJSON();
    this.propsStorage.save(key, data);
  }

  saveSheetId(key = this.tocSheetIdKey){
    this.propsStorage.save(key, this.sheetId)
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
    this.saveSheetId();
    return sheet

  }

  fetchSheetId() {
    this.sheetId = this.sheet.getSheetId();
  }

  setName(newName) {
    this.name = newName;
  }

  setContentIds(sheetIds) {
    this.contentIds = sheetIds;
  }

  setAllSheetIds(sheetIds) {
    this.allSheetIds = sheetIds;
  }

  setRangeHeaderA1Notation(str) {
    this.rangeHeaderA1Notation = str;
  }
  setRangeContentsA1Notation(str) {
    this.rangeContentsA1Notation = str;
  }

  fetchSheetIds() {
    return this.spreadsheetUtil.getSheetIds();
  }

  updateSheetIds() {
    this.allSheetIds = this.fetchSheetIds();
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

  toJSON() {
    return {
      name: this.name,
      key: this.key,
      backupKey: this.backupKey,
      tocSheetIdKey: this.tocSheetIdKey,
      sheetId: this.sheetId,
      titles: this.titles,
      allSheetIds: this.allSheetIds,
      contentIds: this.contentIds,
      rangeHeaderName: this.rangeHeaderName,
      rangeHeaderA1Notation: this.rangeHeaderA1Notation,
      rangeContentsName: this.rangeContentsName,
      rangeContentsA1Notation: this.rangeContentsA1Notation,
      dataRangeValues: this.dataRangeValues,

    }
  }

  setNamedRange(range, name) {
    this.spreadsheetUtil.setNamedRange(range, name)
  }


  fetchSheetValues() {
    const sheet = this.spreadsheetUtil.getSheetById(this.sheetId);
    const data = sheet.getDataRange().getValues();
    this.data = data;
  }

  load(key) {
    if (typeof key !== "string") {
      throw new Error("Please provide valid key.");
    }
    return this.propsStorage.load(key);
  }

  loadTocSheetId(){
    const tocSheetId =  this.propsStorage.load(this.tocSheetIdKey);
    if(tocSheetId){
      return tocSheetId;
    }
  }

  getRichTextValues() {
    const rangeContents = this.spreadsheetUtil.getRangeByName(this.rangeContentsName);
    this.richTextValues = rangeContents.getRichTextValues();
  }

  updateBackup(){
    this.backup = this.toJSON();
  }

  saveBackup() {
    // const { name, backupKey, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation } = this;
    // this.backup = { name, backupKey, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation };
    this.backup = this.toJSON();
    this.save(this.backupKey, this.backup);
  }

  getBackUp() {

    const data = this.toJSON();
    return data;
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
    return this.spreadsheetUtil.getSheetById(this.loadTocSheetId())
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
    let sheet;
    try {
      range = range || this.getRangeByName(this.rangeHeaderName);
    } catch (getError) {
      console.error(getError)
    }

    sheet = this.sheet;
    if (!sheet || (sheet && TocSheet.isEmptyObject(sheet))) {
      try {
        sheet = this.fetchSheet();
      } catch (fetchError) {
        console.error("Could not retrieve the sheet: ", error)
        return; //Exit, cannot getRange to setvalues
      }
    }

    range = sheet.getRange(this.rangeHeaderA1Notation);
    if (range) {
      this.setNamedRange(this.rangeHeaderName, range)
      range = this.getRangeByName(this.rangeHeaderName);
      if (range) {
        range
          .setValue(this.name)
          .setFontWeight("bold");
      } else {
        console.error("Could net set values for the header.  Check the header range name: ", this.rangeHeaderName);
      }
    } else {
      console.error("Could not retrieve the range.  Check the header range A1 Notation: ", this.rangeHeaderA1Notation);
    }
  }


  pasteSheetLinks(range, links = this.links) {
    // Ensure links is an array with elements
    links = (links && links.length) ? links : this.createSheetLinks();

    try {
      range = range || this.getRangeByName(this.rangeContentsName);
    } catch (err) {
      console.error("Could not get range by name: ", err)
    }

    let sheet = this.sheet;
    if (!sheet || (sheet && TocSheet.isEmptyObject(sheet))) {
      try {
        sheet = this.fetchSheet();
      } catch (fetchErr) {
        console.error("Could not retrieve the sheet: ", fetchErr)
        return;
      }
    }

    range = sheet.getRange(this.rangeContentsA1Notation);
    if (range) {
      this.setNamedRange(this.rangeContentsName, range);
      range = this.getRangeByName(this.rangeContentsName);
      if (range) {
        range.setRichTextValues(links)
      } else {
        console.error("Could not get the TOC contents range by name.  Check the name set for the TOC contents range: ", this.rangeContentsName)
      }
    } else {
      console.error("Could not get the range.  Check TOC contents range A1 notation: ", this.rangeContentsA1Notation);
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
    const propKeys = [this.key, this.backupKey, this.tocSheetIdKey];
    const sheet = this.fetchSheet();
    if (sheet) {
      this.spreadsheetUtil.deleteSheet(sheet);
    }
    for(let key of propKeys){
      this.propsStorage.deleteSheetProp(key);
    }
  }

  handleMenuSelectRemove(callback = null) {
    this.remove();
    callback();
  }

 
  updateState(newState = {}) {
    for (let key in newState) {
      if (this.hasOwnProperty(key)) {
        this[key] = newState[key]
      } else {
        throw new Error(`${key} is not a property of ${this.name}`)
      }
    }
  }

  restore(backup) {
    console.log(backup)
    if (backup && !TocSheet.isEmptyObject(backup)) {
      const { rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation } = backup
      this.updateState({ rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation })
      this.name = backup.name;
      let newSheet;
      if (this.dataRangeValues) {
        const data = this.dataRangeValues
        newSheet = this.createSheet(backup.name);
        newSheet.getRange(1, 1, data.length, data[0].length).setValues(data)
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