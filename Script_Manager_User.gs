/*  Copyright Alan J. Wells 2020
All rights reserved
This code licensed to be used with the Google Sheets add-on named Script Manager
*/

/*
  This code works together with the Script Manager add-on for Google Sheets -
  The set-up instructions are provided in the add-on and at the GitHub repository

*/

var APP_SETS_SH,GLBLS_O,ADDRESSES;//Global variables - Accessible to all functions

function setGlobals(k) {
try{
  var i,L,name,setsFromSh,value;
  
  try{
    if (GLBLS_O) {return;}//The global variables have already been assigned values
    //Even though the global variables are defined in the global scope they are initially given a value of undefined
    //and an undefined value will not pass the above test
  }catch(e){}
  
  GLBLS_O = {};//Assign an empty object
  ADDRESSES = {};
  
  SS_GLBL = SpreadsheetApp.getActiveSpreadsheet();//Get the spreadsheet that this code is bound to
  APP_SETS_SH = SS_GLBL.getSheetByName('App Settings');//Get the sheet tab named App Settings

  setsFromSh = APP_SETS_SH.getRange(3,1,APP_SETS_SH.getLastRow()-3,2).getValues();
  L = setsFromSh.length;//The number of rows of data that were retrieved
  
  for (i=0;i<L;i++) {//Process every element of data
    name = setsFromSh[i][0];//Get the name of the setting
    
    if (name == "") {//If there is a blank row
      continue;//then dont process this data element
    };
    
    value = setsFromSh[i][1];//The setting value
    GLBLS_O[name] = value;//Add the value to the global object
    ADDRESSES[name] = "B" + (i+3).toString();
  }
}catch(e){
  //lts('Error 30',e.message)
  errorHandler({msg:e.message,stack:e.stack});
} 
}

setGlobals();//Run the fnc to assign values to the global variables

function OnChangeInUserCode(e) {//This is the fnc that will be triggered by the add-on after you install the trigger
try{
  var action,arrActions,fileContent,hadError,i,L,responseCode,sourceContent,success,theseActions;
  /*
  This function is the function that runs from the ON_CHANGE event -
  You need to install a trigger -
  The trigger only needs to be installed once -
  You must click the Edit menu and then click Current project's triggers and a new browser tab will open -
  Add a new trigger - Click the button to Add Trigger -
  Choose which function to run = OnChangeInUserCode
  Select event type = On change
  Scroll to Save button - Click Save
  input parameters can not be directly passed to this code so they must be looked up in the sheet -
  */
  
  //lts('testOnChange user code 2','ran')
  //lts('e.changeType user code 3',e.changeType)
  
  if (e.changeType !== 'EDIT') {//The Sheets API causes an ONCHANGE event object but the change type
    //is an EDIT 
    return;
  }
  
  setGlobals();//Make sure that the globals are set

  //lts('APP_SETS_SH.getRange(ADDRESSES.CURRENT_STATUS).getValue() 20',APP_SETS_SH.getRange(ADDRESSES.CURRENT_STATUS).getValue())
  
  if (APP_SETS_SH.getRange(ADDRESSES.CURRENT_STATUS).getValue() === GLBLS_O.DONE_TXT) {
    //lts('they are','equal')
    return;
  }
  
  APP_SETS_SH.getRange(ADDRESSES.CURRENT_STATUS).setValue('RUNNING USER CODE');//Set message to user
  //lts('GLBLS_O.ACTION_SETS 27',GLBLS_O.ACTION_SETS)
  //lts('typeof GLBLS_O.ACTION_SETS 28',typeof GLBLS_O.ACTION_SETS)
  
  if (typeof GLBLS_O.ACTION_SETS === 'string') {
    GLBLS_O.ACTION_SETS = JSON.parse(GLBLS_O.ACTION_SETS);
  }
  
  //lts('typeof GLBLS_O.ACTION_SETS 33',typeof GLBLS_O.ACTION_SETS)
  //lts('GLBLS_O.ACTION_SETS.action 34',GLBLS_O.ACTION_SETS.action)
  //lts('GLBLS_O.ACTION_SETS.settings 35',GLBLS_O.ACTION_SETS.settings)
  
  arrActions = GLBLS_O.ACTION_SETS.settings;
  action = GLBLS_O.ACTION_SETS.action;
  //lts('action 40',action)
  L = arrActions.length;
  //lts('L 45',L)
  
  switch(action) {
    case 'get':
      for (i=0;i<L;i++){
        theseActions = arrActions[i];
        //lts('theseActions 45',theseActions)
        //lts('typeof theseActions 46',typeof theseActions)
        //lts('theseActions.scriptId 47',theseActions.scriptId)
        
        fileContent = appsScriptFileContent({action:'get',scriptId:theseActions.scriptId});
        if (!fileContent) {throw new Error('Cant get file content');}
        
        //lts('typeof fileContent 55',typeof fileContent)
        
        //lts('fileContent 57',fileContent.slice(0,300))
        //lts('arrActions.shNameToWriteTo 58',theseActions.shNameToWriteTo)
        
        success = putContentIntoSh({shTabName:theseActions.shNameToWriteTo,content:fileContent});
        if (success === false) {
          hadError = true;
          break;
        }
      }
      
      break;
    case 'update':
      for (i=0;i<L;i++){
        theseActions = arrActions[i];
        //lts('theseActions 69',theseActions)
        //lts('typeof theseActions 70',typeof theseActions)
        //lts('theseActions.scriptId 71',theseActions.scriptId)
        
        sourceContent = getContentsFromShTab_({shTabName:theseActions.srcShTabName});//Get content from sheet tab
        //lts('sourceContent 75',sourceContent)
        //lts('typeof sourceContent 75',typeof sourceContent)
        
        if (!sourceContent) {throw new Error('Cant get source file content');}
        responseCode = appsScriptFileContent({action:'update',scriptId:theseActions.scriptId},sourceContent);
        
        if (responseCode !== 200) {
          hadError = true;
          errorHandler({msg:'File Update Failed',responseCode:responseCode});
          APP_SETS_SH.getRange(ADDRESSES.ERROR_MESSAGE).setValue('Update Failed - Response Code: ' + responseCode);
        }
      }
      break;
    default:
      throw new Error("Settings are missing for fnc OnChangeInUserCode");
      break;
  }

  if (!hadError) {//There was NOT an error
    APP_SETS_SH.getRange(ADDRESSES.CURRENT_STATUS).setValue(GLBLS_O.DONE_TXT);
  }
  
}catch(e){
  //lts('Error 30',e.message)
  errorHandler({msg:e.message,stack:e.stack});
}
}

function putContentIntoSh(po) {
try{
  var arry,cnt,cutOff,dxEnd,dxStart,increment,L,sh,someContent;
  /*
    po.content - file content as string
    po.shTabName - name of sheet tab to write file content to
  */
  //lts('po 8',po)
  
  if (!po.shTabName) {
    throw new Error('po.shTabName not passed into fnc putContentIntoSh');
  }
  
  //lts('typeof po.content 15',typeof po.content)
  
  if (typeof po.content === 'object') {
    po.content = JSON.stringify(po.content);
  }
  
  increment = 44444;
  arry = [];
  L = po.content.length;
  //lts('L 18',L)
  
  sh = SS_GLBL.getSheetByName(po.shTabName);
  sh.deleteRows(1, sh.getLastRow());
  
  cnt = 0;
  cutOff = Math.ceil(L / increment) + 2;//Pad by 2 just to make sure the loop doesnt cut short
  //lts('cutOff 22',cutOff)
  
  dxStart = 0;
  dxEnd = L < increment ? L : increment;
  
  while (cnt < cutOff) {
    cnt++;
    
    someContent = po.content.slice(dxStart,dxEnd);
    //lts('someContent 31',someContent)
    //lts('typeof someContent 32',typeof someContent)
    
    if (!someContent) {
      break;
    }
    arry.push([someContent]);
    dxEnd+=increment;
    dxStart+=increment;
    if (dxStart > L) {break;}//The start index number is greater than the length of the content
  }
  
  //lts('arry 43',arry)
  
  if (sh.getLastRow() < arry.length) {//If the data to set in the sheet has more rows than are in the sheet
    sh.insertRows(1, sh.getLastRow());//Insert some more rows
  }
  
  sh.getRange(1,1,arry.length,1).setValues(arry);
  
  return true;
}catch(e){
  errorHandler({msg:e.message,stack:e.stack});
  return false;
} 
}

function errorHandler(po) {
  var sh;
  
  sh = SS_GLBL.getSheetByName('Log');
  sh.appendRow([new Date(),po.msg,po.responseCode,po.stack]);
  try{
    APP_SETS_SH.getRange(ADDRESSES.CURRENT_STATUS).setValue('ERROR');
  }catch(e){
  
  }
}

function getContentsFromShTab_(po) {
try{
  var allContent,arry,content,i,L,rslt,sh,someContent;
  /*
    po.shTabName - the name of the sheet tab to get the file content out of 
  */
  
  //lts('po 51',po)
  content = "";
  
  sh = SS_GLBL.getSheetByName(po.shTabName);
  
  allContent = sh.getRange(1,1,sh.getLastRow(),1).getValues();
  //lts('allContent 51',allContent)
  
  L = allContent.length;
  
  for (i=0;i<L;i++){
    someContent = allContent[i];
    content += someContent;
  }

  //lts('content 94',content)
  
  return content;
}catch(e){
  errorHandler({msg:e.message,stack:e.stack});
}
}

function testWriteTofile() {
  
  putContentIntoSh({shTabName:toDo,content:"Test Content"});
  
  
}

function appsScriptFileContent(po,content) {
try{
  var accessTkn,fileContent,fileID,mediaData,method,options,payload,response,resource,url;

  /*
    po.action - update or get - required - update the Apps Script file orget file content
    po.scriptId - The file ID of the Apps Script file - required
    content - the Apps Script file content - required IF the action is update
  */
  
  //lts('po 15',po)
  //lts('po.scriptId 16',po.scriptId )
  //lts('content 17',content)

  if (!content && po.action !== 'get') {
    throw new Error('The file content was not passed into the fnc');
    return;
  }

  if (typeof content === 'object') {
    content = JSON.stringify(content);
  }
  
  accessTkn = ScriptApp.getOAuthToken();
  
  url = "https://script.googleapis.com/v1/projects/" + po.scriptId + "/content";
  //lts('url 26',url)
  
  method = po.action === 'update' ? 'PUT' : 'GET';
  //lts('method 29',method)

  options = {
    "method" : method,
    "muteHttpExceptions": true,
    "headers": {
      'Authorization': 'Bearer ' +  accessTkn
     }
  };

  if (po.action === 'update') {
    options.contentType = "application/json";//If the content type is set then you can stringify the payload
    options.payload = content;//stringified the content  
  }
  
  response = UrlFetchApp.fetch(url,options);//Make an external request and get a response back

  //lts('typeof response',typeof response)
  //lts('response 39',JSON.stringify(response))
  //lts('response.getResponseCode 49',response.getResponseCode())
  //lts('response.getContentText 50',response.getContentText())

  if (response.getResponseCode() !== 200) {
    throw new Error('Error getting file: ' + response.error);
  }
  
  if (po.action === 'get') {
    fileContent = response.getContentText();
    fileContent = JSON.parse(fileContent);
    fileContent = fileContent.files;//Only get the files
    //lts('fileContent 60',fileContent.slice(0,200))
    return fileContent;
  } else {//The file was updated
    return response.getResponseCode();
  }
  
} catch(e) {
  //ll(response)
  errorHandler({msg:e.message,stack:e.stack});
  return null;
}
};

function lts(a,b) {
  var SH_ERROR = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  
  b = processValForLog(b);
  
  SH_ERROR.appendRow([new Date(),a,b,'User Code']);
}

function processValForLog(val) {
try{
  var typeOfThisVal;
  
  typeOfThisVal = typeof val;

  if (val) {
    if (typeOfThisVal === 'object') {//The variable val can be an object and still be null
      
      if (Array.isArray(val)) {//The very first test should be for an array because that is easy to check for
        //console.log('it is an array')
        if (val.toString().indexOf("[object Object]") !== -1) {
          val = JSON.stringify(val);
        } else {
          val = val.toString();
        }
        //Logger.log('val: ' + val)
      } else {//It is an object but not an array
        //Logger.log('its NOT an array')
        try{
          val = JSON.stringify(val);//Test for whether it is a date
        }catch(e) {//If this is an invalid date then JSON stringify will fail
          //console.log('Error stringifying')
        }
      }
      
      //Logger.log('typeof val: ' + typeof val)
      
        if (typeof val !== 'string') {
          try{
            val = val.toString();
          }catch(e) {
            val = e.message;
          }
        }
      
      if (val.indexOf("{") !== 0 && typeof val !== 'string') {
        val = val.toString();
      }
      
      //continue;
    }
  }
  
  if (typeOfThisVal === 'number') {
    val = '"' + val.toString();
  }
  
  if (val === undefined) {//Avoid having all three cells undefined
    val = 'UNDEFINEDDD';
  } else if (val === null) {
    val = 'NULLLL';
  } else if (val === false) {
    val = 'FALZZE';
  } else if (val === true) {
    val = 'TREWWW';
  }
  
  return val;
}catch(e) {
  var f = new Function('e','Logger.log(e.message + ":\n\n" + e.stack)');
  f(e);
}
}


