
/**
 * /////////////CLOSURE LOGIC THE CREATES AN INSTANCE OF THE SERVICES TO USE IN THE GLOBAL SPACE///////////
 * The nested functions "stores" the data from the outter function
 * The returned nested functions holds this data to be stored in a global variable for reuse
 * Think of the outer function as the parent and the inner function as the child
 */

/////////GLOBAL SERVICES STORED ACCESS THROUGH THE CLOSURE FUNCTION//////////////
const getSpreadsheetApp = getService('SpreadsheetApp');
const getPropsService = getService('PropertiesService');
const getScriptApp = getService("ScriptApp")

// ... use getSpreadsheetApp and getUrlFetchApp as needed



///////////////CLOSURE FUNCTION THAT CREATES ONE GLOBAL INSTANCE OF THE SERVICES USED IN THIS PROJECT/////////////
function getService(serviceName) {
  const googleServices = {
    "SpreadsheetApp": SpreadsheetApp,
    "PropertiesService": PropertiesService,
    "ScriptApp": ScriptApp
  }
  const services = {}; // Internal storage for service instances;
  if(googleServices.hasOwnProperty(serviceName)){
    return function() {
      if (!services[serviceName]) {
        // Use a generic approach to get the service instance
        services[serviceName] = googleServices[serviceName];
      }
      return services[serviceName];
    };
  }
  return undefined;
}




