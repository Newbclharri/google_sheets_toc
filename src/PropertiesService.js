class PropertiesServiceStorage {

  constructor(){
    this.key = "tocSheet";
  }

  save(tocData){
    const data = JSON.stringify(tocData)
    return PropertiesService.getScriptProperties().setProperty(this.key, data)
  };

  load(){
    const data = PropertiesService.getScriptProperties().getProperty(this.key)
    if(!data){
      return null
    };
    return JSON.parse(data);
  }  
}
