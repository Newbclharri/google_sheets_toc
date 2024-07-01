class TocSheet {

  constructor(params = {}, spreadsheetUtil, propsStorage) {
    this.sheetId = params.sheetId || null;
    this.allSheetIds = params.allSheetIds || null;
    this.ssUtil = spreadsheetUtil || SpreadsheetUtility.getInstance();
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

  static convertToNumber(value) {

    //the passed argument is already a number:
    if (typeof value === "number") {
      return value;
    }


    //the passed argument is not a number, but also not a string
    if (typeof value !== "string") {
      throw new Error(`Expected string. Received ${value}: ${typeof value}`)
    }

    const convertedValue = Number(value);
    if (isNaN(convertedValue)) {
      throw new Error(`Conversion to number failed for value: ${value}`);
    }

    return convertedValue;

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

  saveSheetId(key = this.tocSheetIdKey) {
    this.propsStorage.save(key, this.sheetId)
  }

  initialize(name = this.name) {
    this.createSheet(name);
    this.allSheetIds = this.fetchSheetIds();
    this.contentIds = this.fetchSheetIdsNotEqualTo(this.sheetId)
    this.titles = this.fetchTitleNames();
    this.createSheetLinks(this.contentIds);
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
    const sheet = this.ssUtil.insertSheet(this.name)
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

  setTitles(titles) {
    if (!titles.length) {
      console.log("No titles to update.");
      return;
    }
    this.titles = titles;
    console.log(`${titles.length} titles added successfully: ${titles}`)
    return;
  }

  updateTitlesByIds(ids) {
    if (!ids.length) {
      console.log("No ids found to update titles.");
      return;
    }

    //get the sheet names by id
    const names = [];
    ids.forEach(id => {
      if (!isNaN(id)) {
        names.push(this
          .getSheetById(id)
          .getName())
      }
    });

    //push new titles to the titles array;
    try {
      if (Array.isArray(this.titles)) {
        names.forEach(name => this.titles.push(name));
        console.log(`Successfully added ${names.length} titles: ${names}`);
      } else {
        throw new Error("Could not find an existing array of titles.")
      }
    } catch (err) {
      console.error("New title array creatad: ", err.stack)
      this.titles = this.fetchTitleNames();
    }
  }

  setContentIds(sheetIds) {
    this.contentIds = sheetIds;
    console.log(`ContentIds: ${this.contentIds}`)
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
    return this.ssUtil.getSheetIds();
  }

  updateSheetIds() {
    this.allSheetIds = this.fetchSheetIds();
  }


  fetchSheetIdsNotEqualTo(id = this.sheetId) {
    return this.ssUtil.getSheetIdsNotEqualTo(id);
  }

  createSheetLinks(contentIds = this.contentIds) {
    if (!(contentIds && contentIds.length)) {
      throw new Error("No ids to convert to links")
    }

    const links = [];
    let link;
    contentIds.forEach(id => {
      link = this.ssUtil.createSheetLink(id, false, true)
      links.push([link])
    });

    return links;
  }


  fetchTitleNames() {
    let titles;
    try {
      const sheets = this.ssUtil.getSheets();
      const sheetId = this.sheetId;
      titles = sheets.filter(sheet => sheet.getSheetId() !== sheetId)
        .map(sheet => sheet.getName());
    } catch (err) {
      console.err(err.message);
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

  getContentIds() {
    return this.contentIds || this.fetchContentIds();
  }

  updateContentIds(newIds) {
    const currContentIds = this.getContentIds()
    if (currContentIds) {
      const count = 0;
      const added = []
      newIds.forEach(id => {
        if (!currContentIds.includes(id)) {
          currContentIds.push(id);
          added.push(id)
          count++;
        }
      })
      console.log(`Added ${count} content ids: ${added} \nAll content Ids: ${this.contentIds}`);
    }
  }

  getSheetById(id) {
    const sheet = this.ssUtil.getSheetById(id);
    return sheet;
  }

  fetchContentIds() {
    try {
      const contentIds = this.fetchSheetIdsNotEqualTo(this.sheetId);
      if (!contentIds) {
        throw new Error("No ids to convert to links");
      }
    } catch (err) {
      console.error(err.stack);
    }
  }

  setNamedRange(range, name) {
    this.ssUtil.setNamedRange(range, name)
  }


  // getRangeContents() {
  //   let rangeHeader, rangeContents;

  //   try {

  //     // Fetch the sheet and attempt to get the range by A1 notation
  //     const sheet = this.fetchSheet();
  //     if (!sheet) {
  //       throw new Error("Sheet not found.");
  //     }
  //     const lastRow = sheet.getLastRow();
  //     rangeHeader = this.getRangeByName(this.rangeHeaderName) || sheet.getRange(1,1)

  //     rangeContents = this.getRangeByName(this.rangeContentsName);
  //     if (!rangeContents) {
  //       //Attempt to get range contents via A1Notation
  //       rangeContents = sheet.getRange(this.rangeContentsA1Notation);
  //     }

  //     if (!rangeContents) {
  //       //set range to default (1st row, 1st column);
  //       const rowsHeader = rangeHeader.getLastRow();
  //       const startRow = rowsHeader + 1;
  //       rangeContents = sheet.getRange(startRow, 1, lastRow - rowsHeader);
  //     }

  //     //adjust range for data processing
  //     const rowsHeader = rangeHeader.getLastRow();
  //     const startRow = rowsHeader + 1;
  //     const startColumn = rangeContents.getColumn();
  //     const adjustedRange = sheet.getRange(startRow, startColumn, lastRow - rowsHeader);

  //     //Reset the contents named range
  //     this.setNamedRange(this.rangeContentsName, adjustedRange);

  //     //Update contents A1 notation
  //     this.rangeContentsA1Notation = this.getRangeByName(this.rangeContentsName).getA1Notation;
  //     return adjustedRange;

  //   } catch (err) {
  //     console.error(err);     
  //   }
  //   // If no range is found or set, return null or handle as needed
  //   return null;
  // }

  getRangeContents() {
    let rangeHeader, rangeContents;
  
    try {
      // Fetch the sheet and attempt to get the range by A1 notation
      const sheet = this.fetchSheet();
      if (!sheet) {
        throw new Error("Sheet not found.");
      }
  
      const lastRow = sheet.getLastRow();
      rangeHeader = this.getRangeByName(this.rangeHeaderName) || sheet.getRange(1, 1);
  
      rangeContents = this.getRangeByName(this.rangeContentsName);
      if (!rangeContents) {
        // Attempt to get range contents via A1 notation
        rangeContents = sheet.getRange(this.rangeContentsA1Notation);
      }
  
      if (!rangeContents) {
        // Set range to default (1st row, 1st column)
        const rowsHeader = rangeHeader.getLastRow();
        const startRow = rowsHeader + 1;
        rangeContents = sheet.getRange(startRow, 1, lastRow - rowsHeader);
      }
  
      // Get the values in the range and find the rows to delete
      const data = rangeContents.getValues();
      const rowsToDelete = [];
      data.forEach((row, index) => {
        //the retrieved range is only one column, so the check for every cell will only be one column in length
        // ex. [[cell1], [cell1], etc]
        if (row.every(cell => cell === "")) {
          rowsToDelete.push(index + rangeContents.getRow());
        }
      });
  
      // Remove rows from bottom to top to avoid index shifting issues
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
      }
  
      // Recalculate the range after removing blank rows
      const newLastRow = sheet.getLastRow();
      const rowsHeader = rangeHeader.getLastRow();
      const startRow = rowsHeader + 1;
      const startColumn = rangeContents.getColumn();
      const adjustedRange = sheet.getRange(startRow, startColumn, newLastRow - rowsHeader);
  
      // Reset the contents named range
      this.setNamedRange(this.rangeContentsName, adjustedRange);
  
      // Update contents A1 notation
      this.rangeContentsA1Notation = this.getRangeByName(this.rangeContentsName).getA1Notation();
      return adjustedRange;
  
    } catch (err) {
      console.error(err);
    }
  
    // If no range is found or set, return null or handle as needed
    return null;
  }
  


  fetchSheetValues() {
    const sheet = this.ssUtil.getSheetById(this.sheetId);
    const data = sheet.getDataRange().getValues();
    this.data = data;
  }

  load(key) {
    if (typeof key !== "string") {
      throw new Error("Please provide valid key.");
    }
    return this.propsStorage.load(key);
  }

  loadTocSheetId() {
    const tocSheetId = this.propsStorage.load(this.tocSheetIdKey);
    if (tocSheetId) {
      return tocSheetId;
    }
  }

  getRichTextValues() {
    const rangeContents = this.ssUtil.getRangeByName(this.rangeContentsName);
    this.richTextValues = rangeContents.getRichTextValues();
  }

  updateBackup() {
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
      console.err(err.message);
      return undefined;
    }
  }

  fetchSheet() {
    try {
      const id = this.loadTocSheetId();
      if (!id) {
        throw new Error("The sheet ID is null or undefined.");
      }

      const sheet = this.ssUtil.getSheetById(id);
      if (!sheet) {
        throw new Error(`No sheet found with ID: ${id}`);
      }

      return sheet;

    } catch (err) {
      console.error("Error fetching the sheet", err.stack);
      return undefined;
    }
  }

  doesExistSheet() {
    const sheet = this.fetchSheet()
    return sheet ? true : false;
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
      console.err(err.message);
      return this.allSheetIds = []
    }
  }



  getRangeByName(name) {
    const range = this.ssUtil.getRangeByName(name)

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
      console.err(err, `| Could not find sheet with id . ${this.sheetId}`)
    }
  }
  remove() {
    const propKeys = [this.key, this.backupKey, this.tocSheetIdKey];
    const sheet = this.fetchSheet();
    if (sheet) {
      this.ssUtil.deleteSheet(sheet);
    }
    for (let key of propKeys) {
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
      console.log(`New Sheet Id: ${this.sheetId}`);
      //const {name, titles, sheetId, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation} = this;
      //dataToSave = {name, titles, sheetId, allSheetIds, rangeHeaderName, rangeHeaderA1Notation, rangeContentsName, rangeContentsA1Notation}
      //destructure if more efficient before saving
      //const {this object property names here} = this;
      //this.save("tocSheet", this);
    }
  }
}