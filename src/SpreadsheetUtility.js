class SpreadsheetUtility {
  constructor() {
    if (SpreadsheetUtility.instance) {
      return SpreadsheetUtility.instance
    }
    this.spreadsheetApp = getSpreadsheetApp(); //|| SpreadsheetApp;
    this.spreadsheet = this.spreadsheetApp.getActive();
    this.sheets = this.spreadsheet.getSheets();
    this.newTextStyle = this.spreadsheetApp.newTextStyle();
    this.newRichTextValue = this.spreadsheetApp.newRichTextValue();
    SpreadsheetUtility.instance = this;
  }

  static getInstance() {
    if (!SpreadsheetUtility.instance) {
      return new SpreadsheetUtility();
    }
    return SpreadsheetUtility.instance;
  }

  getActive() {
    return this.spreadsheet;
  }

  getUi() {
    this.spreadsheetApp.getUi();
  }

  insertSheet(name) {
    return this.spreadsheet.insertSheet(name, 0);

  }

  setActiveSheet(sheet) {
    return this.spreadsheet.setActiveSheet(sheet);
  }

  setNamedRange(name, range) {
    this.spreadsheet.setNamedRange(name, range)
  }

  getSheets() {
    return this.sheets;
  }

  getNumSheets() {
    return this.sheets.length;
  }

  getSheetByName(name) {
    return this.spreadsheet.getSheetByName(name)
  }
  getSheetById(id) {
    if (isNaN(id)) {
      throw new Error(`${id} is not a number.`)
    }

    try {
      const sheet = this.sheets.find(sheet => sheet.getSheetId() == id);
      if (!sheet) {
        throw new Error(`Could not find with id ${id}`);
      }
      return sheet
    } catch (err) {
      console.error(err.stack);
    }
  }

  getSheetIds() {
    return this.spreadsheet.getSheets().map(sheet => sheet.getSheetId());
  }

  getSheetIdsNotEqualTo(tocId) {
    const ids = this.sheets.filter(sheet => sheet.getSheetId() !== tocId)
      .map(sheet => sheet.getSheetId());
    return ids

  }

  getRangeByName(name) {
    return this.spreadsheet.getRangeByName(name)
  }

  getA1Notation(range) {
    return range.getA1Notation()
  }

  createSheetLink(sheetId, url = null, underline = false, bold = false) {
    if (!isNaN(sheetId)) {
      const sheet = this.getSheetById(sheetId);
      const sheetUrl = `#gid=${sheetId}`;
      const linkStyle = this.newTextStyle
        .setUnderline(underline)
        .setBold(bold)
        .build();
      const link = this.newRichTextValue
        .setText(sheet.getName())
        .setLinkUrl(sheetUrl)
        .setTextStyle(linkStyle)
        .build()
      return link;
    }
  }

  createSheetLinks(sheetIds, underline = false, bold = false) {
    //get all sheets from passed sheetIds
    const sheets = [];
    const links = [];
    let sheetUrl, linkStyle, link;
    sheetIds.forEach(id => {
      const sheet = this.getSheetById(id);
      if (sheet) {
        link = this.createSheetLink(sheet, underline, bold);
        links.push([link]);
      }
    })
    return links;
  }

  deleteSheet(sheet) {
    this.spreadsheet.deleteSheet(sheet);
  }
}