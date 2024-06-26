class TriggerManager{
  constructor(scriptApp, spreadsheetUtil){
    if(TriggerManager.instance){
      return TriggerManager.instance;
    }
    this.scriptApp = scriptApp; //|| ScriptApp;
    this.projectTriggers = this.scriptApp.getProjectTriggers(); 
    this.docType = spreadsheetUtil.getActive(); // || SpreadsheetApp.getActive();
    this.ON_CHANGE = this.scriptApp.EventType.ON_CHANGE;
    this.ON_EDIT = this.scriptApp.EventType.ON_EDIT;
    TriggerManager.instance = this;
  }
  //additional methods below:
  static getInstance(scriptApp, spreadsheetUtil){
    if(!TriggerManager.instance){
      scriptApp = scriptApp; //|| ScriptApp;
      spreadsheetUtil = spreadsheetUtil; //|| SpreadsheetApp;
      return new TriggerManager(scriptApp, spreadsheetUtil)
    }
  }
  hasTrigger(eventType, handlerFunction){
    return this.projectTriggers.some(trigger => 
      (trigger.getEventType() === eventType && trigger.getHandlerFunction() === handlerFunction)
    );
  };

  getEventType(){
    return{
      ON_CHANGE: this.ON_CHANGE,
      ON_EDIT: this.ON_EDIT
    }
  }

  setTrigger(eventType, handlerFunction){
    if(!this.hasTrigger(eventType, handlerFunction)){
      switch(eventType){
        case this.ON_CHANGE:
          this.scriptApp.newTrigger(handlerFunction)
          .forSpreadsheet(this.docType)
          .onChange()
          .create();
        break;
        //add additional eventType cases if needed
        default:
          console.log(`Error, could not set trigger. Event Type: '${eventType}' Handler Function: '${handlerFunction}'`);
      }
      return console.log(`Trigger with Event Type: '${eventType}' and Handler Function: '${handlerFunction}' set!`);
    }
    return console.log(`Trigger with Event Type: '${eventType}' and Handler Function: '${handlerFunction}' already exists. No trigger was set.`);
  }

  deleteTrigger(eventType, handlerFunction){
    if(this.hasTrigger(eventType, handlerFunction)){
      const trigger = this.projectTriggers.find(trigger => 
        trigger.getEventType() === eventType && trigger.getHandlerFunction() === handlerFunction);
        this.scriptApp.deleteTrigger(trigger);
        return console.log(`Removed trigger with Event Type: '${eventType}' and Handler Function: '${handlerFunction}'.`);
    }
    return console.log(`Could not find trigger with Event Type: '${eventType}' and Handler Function: '${handlerFunction}'. 
      No trigger removed.`);
  }
}

function testSetTrigger(){
  const triggerManager = new TriggerManager();
  const eventType = triggerManager.getEventType().ON_CHANGE;
  const handlerFunction = "handleOnChange"
  triggerManager.setTrigger(eventType, handlerFunction);
}