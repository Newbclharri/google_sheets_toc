class PropertiesServiceStorage {

  constructor(){
    this.key = "tocSheet";
    this.scriptProps = PropertiesService.getScriptProperties();
  }

  save(tocData){
    const data = JSON.stringify(tocData)
    return this.scriptProps.setProperty(this.key, data);
  };

  load(){
    const data = this.scriptProps.getProperty(this.key);
    return JSON.parse(data);
  }
  
  deleteSheetProp(key){
    return this.scriptProps.deleteProperty(key);
  }
}
