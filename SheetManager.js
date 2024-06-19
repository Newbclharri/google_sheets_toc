/**
 * Class representing a Table of Contents Manager.
 */
class SheetManager {
  /**
   * Create a Table of Contents Manager.
   */
  constructor(name, spreadsheetApp, propsService,scriptApp) {
    if(SheetManager.instance){
      return SheetManager.instance;
    }

    this.name = name;
    /**
     * The PropertiesService for managing script properties.
     * @type {Properties}
     */
    this.spreadsheet = spreadsheetApp();
    this.propsService = propsService().getScriptProperties();
    this.scriptApp = scriptApp();
    this.state = {};

    // this.rangeHeaderName = "TOCHeader";
    // this.rangeContentsName = "TOC";
    this.initializeProperties();
    this.createSheet();
    SheetManager.instance = this;
  }


  static getInstance(){
    if(!SheetManager.instance){
      SheetManager.instance = new SheetManager();
    }
    return SheetManager.instance;
  }

  
  /**
   * Initialize the manager with properties from the PropertiesService.
   */

  initializeProperties(){
    const name = this.name.toLowerCase().replace(/\s+/g,'')
    const sheetIdKey = name + "Id";
    console.log(sheetIdKey)
    const propNames = [sheetIdKey];
    const scriptProps = this.getScriptProps()
    propNames.forEach(prop =>{
      if(scriptProps.hasOwnProperty(prop)){
        this[prop] = scriptProps[prop]
      }
    })

    //set state variables
    this.state.sheetExists = false;
    this.state.doRename = false;
    this.state.doDelete = false;
  }

  createSheet(){
    if(! JSON.parse(this.getScriptProp("sheet"))){
      const name = this.name.toLowerCase().replace(/\s+/g,'')
      const sheetIdKey = name + "Id";
      this.sheet = this.spreadsheet.getActive().insertSheet(this.name, 0)
      this[sheetIdKey] = this.sheet.getSheetId();
      this.state.sheetExists = true;
      this.setScriptProp("sheet", this.sheet)
      return this.sheet;
    }
  }

  getSheet(){
    return this.sheet;
  }

  getSheetId(){
    return this.sheet.getSheetId();
  }

  getPropsInstance(){
    const propsInstance = this.getScriptProp(JSON.parse("sheet"))
    return propsInstance
  }

  /**
   * Set a single property.
   * @param {string} key - The key of the property.
   * @param {string} value - The value of the property.
   */
  setScriptProp(key, value) {
    this[key] = value;
    this.propsService.setProperty(key, value);
  }

  /**
   * Set multiple properties.
   * @param {Object} props - An object containing key-value pairs of properties.
   */
  setScriptProps(props) {
    for (let key in props) {
      this[key] = props[key];
    }
    this.propsService.setProperties(props);
  }

  setState(key, value){
    this.state[key] = value;
  }

 


  /**
   * Get a single property.
   * @param {string} prop - The key of the property.
   * @returns {string} The value of the property.
   */
  getScriptProp(prop) {
    return this.propsService.getProperty(prop)
  }

  /**
   * Get all properties.
   * @returns {Object} An object containing all properties.
   */
  getScriptProps() {
    return this.propsService.getProperties();
  }

  deleteProp(prop){
    this.propsService.deleteProperty(prop)
  }

  start() {
    this.setProps(this.propsService().getProperties());
  }

  runFromMenu(doSort){
    this.setProps({"doSortToc":doSort, "doInitialize":true});
    main(true)
  }

  /**
   * 
   */
  insertSheet(){
    if(!this.getProp("tocSheetId")){
      const sheet = this.ss.insertSheet(this.name,0);
      const sheetId = sheet.getSheetId();
      this.setProp("tocSheetId", sheetId)
      this.setNamedRanges(sheet);
    }else{
      let tocId = this.getProp("tocSheetId")
      tocId = Number(tocId)
      const sheet = this.utility.getSheetById(tocId)
      this.ss.setActiveSheet(sheet)
    }
  }

  /**
   * Set a property to indicate that the Table of Contents should be removed.
   */
  deleteToc() {
    const tocSheetId =  this.propertyToNumber(this.sheetId)
    const sheet = this.utility.getSheetById(tocSheetId);
    if(sheet){
      this.ss.delete(sheet)
    }
    // this.setProp("doRemoveToc", true);
    
  }

  /**
   * Set a trigger for the onChange event if it does not already exist.
   */
  setTriggers() {
    TriggerManager.setTriggers();
  }

}
