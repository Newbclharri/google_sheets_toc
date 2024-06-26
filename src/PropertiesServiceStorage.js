class PropertiesServiceStorage {

  constructor(){
    if(PropertiesServiceStorage.instance){
      return PropertiesServiceStorage.instance
    }
    this.key = "tocSheet";
    this.propsServ = getPropsServ(); // || PropertiesService;
    this.scriptProps = this.propsServ.getScriptProperties();
    PropertiesServiceStorage.instance = this;
  }

  static getInstance(){
    if(!PropertiesServiceStorage.instance){
      return new PropertiesServiceStorage();
    }
    return PropertiesServiceStorage.instance;
  }

  save(key = this.key, tocData){
    const data = JSON.stringify(tocData)
    return this.scriptProps.setProperty(key, data);
  };

  batchSave(obj){
  
    this.scriptProps.setProperties(obj)
  }

  load(key = this.key){
    const data = this.scriptProps.getProperty(key);
    return JSON.parse(data);
  }
  
  deleteSheetProp(key=this.key){
    return this.scriptProps.deleteProperty(key);
  }
}
