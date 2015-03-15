var SPREADSHEET_ID = "ID_OF_SPREADSHEET_WHERE_YOU_WANT_DATA_TO_BE_SAVED"
var AUTHORIZE_URL = 'https://www.linkedin.com/uas/oauth2/authorization';
var TOKEN_URL = 'https://www.linkedin.com/uas/oauth2/accessToken';
var CLIENT_ID = 'FROM_LINKEDIN_APP'; // create a new app here: https://www.linkedin.com/secure/developer?newapp=
var CLIENT_SECRET = 'FROM_LINKEDIN_APP';
var REDIRECT_URL= ScriptApp.getService().getUrl();

//this is the user propety where we'll store the token & state parameter for authorize URL
var tokenPropertyName = 'LINKEDIN_OAUTH_TOKEN'; 
var statePropertyName = 'STATE_VALUE_AUTH'

/*
 * Function called first when Web App exectues
 */
function doGet(e) {    
  UserProperties.setProperty(tokenPropertyName, '');
  if(e.parameters.code){//if we get "code" as a parameter in, then this is a callback. we can make this more explicit
    getAndStoreAccessToken(e.parameters.code);
    Logger.log("state = " + e.parameters.state);
    return HtmlService.createHtmlOutput("<html><h2>"+populateLinkedInData(true, e.parameters.state)+"</h1></html>");
  }
  else if(isTokenValid()){//if we already have a valid token, go off and start working with data        
    return HtmlService.createHtmlOutput("<html><h2>"+populateLinkedInData(false,'')+"</h1></html>");
  }
  else {//we are starting from scratch or resetting
    return HtmlService.createHtmlOutput("<html><h1>Fetch LinkedIn Data</h1><a href='"+getURLForAuthorization()+"'>click here to start</a></html>");
  } 
}

/*
 * Populate LinkedIn Data in "LinkedIn_BuiltWith_Company_Check_Production" Spreadsheet
 */
function populateLinkedInData(authorization, state) {
  // is state valid
  if (authorization)
    if (state != UserProperties.getProperty(statePropertyName))
      return "<h2>Access Denied!</h2>"; 
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);  // production sheet
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var domains = sheet.getRange(3, 1, lastRow-1, 1).getValues();
  var linkedInId = sheet.getRange(3,2, lastRow-1, 1).getValues();
  var token = UserProperties.getProperty(tokenPropertyName); 
  
  //Filter for empty rows
  for(var i = domains.length-1; i>=0; i--){
    if(domains[i][0] != '') {
      lastRow = i+1;
      break;
    }
  }  
  // fetch data here
  var fetchArgs = {method : 'get', muteHttpExceptions: false, "headers" : {"Accept" : "application/xml"}};
  var url = 'https://api.linkedin.com/v1/companies'; //?email-domain=trulia.com
  var reWebAddress = /([A-Za-z0-9][A-Za-z0-9-]{1,62}\.[A-Za-z\.]+)/; // regEx to test if it is web address
  for(i=0; i<lastRow; i++){
    if(linkedInId[i][0] != '') continue; //Leave records which are already processed
    ss.toast('Please standby...Processing row '+(i+2)+' using LinkedIn API');
    var reqUrl;
    if (reWebAddress.test(domains[i][0])) {
        reqUrl = url + '?email-domain='+domains[i][0]+'&oauth2_access_token='+token;
    } else if (isNaN(domains[i][0])) {
      //not a website address, not Company ID
      reqUrl = url + '/universal-name='+domains[i][0]+'?oauth2_access_token='+token;
    }    
    else {      
      // Numerice. Assuming it to be company ID
      reqUrl = url + '/id='+domains[i][0]+'?oauth2_access_token='+token;
    }
    var data;
    var companyXml
    try{ companyXml= UrlFetchApp.fetch(reqUrl, fetchArgs).getContentText(); }
    catch(e){data = [e.message.replace(/[\r\n]/g, '')], companyXml = undefined}
    if(companyXml != undefined){
      ; //if error is returned by API then skip that record
      var obj = getCompanyId_(companyXml);
      
      if(obj == 'NOT FOUND') continue; //if company not found then skip that record
      
      var mainUrl = url+'/'+obj.id+':(id,company-type,industries,ticker,status,locations,email-domains,blog-rss-url,twitter-id,description,founded-year,num-followers,employee-count-range,specialties)'+'?oauth2_access_token='+token;
      var xmlRaw
      try {
       xmlRaw = UrlFetchApp.fetch(mainUrl, fetchArgs).getContentText();
      }
      catch(e2) {
        //reset the token
        UserProperties.setProperty(tokenPropertyName, '');
        return "<h1>Fetch LinkedIn Data</h1><a href='"+getURLForAuthorization()+"'>Throttle limit reached. click here to login from a different user.</a>";     
      }
      data = getDatFromLinkedInXml_(xmlRaw, obj.name);
    }    
    sheet.getRange(i+3, 2, 1, data.length).setValues([data]);    
  }// for loop
    // fetch data over      
  return "<h2>Script Executed Successfully!</h2>";
}

/*
 * Extract data from XML
 */
function getDatFromLinkedInXml_(xmlRaw, companyName){
  
  var company = Xml.parse(xmlRaw).getElement();
  
  var id = company.getElement('id').getText(); //
  
  var name = companyName;
  
  var numFollowers  = company.getElement('num-followers').getText();//
  
  var companyType = '';
  var companyTypeXml  = company.getElement('company-type');
  if(companyTypeXml != null)companyType += companyTypeXml.getElement('name').getText();
  
  var industriesXml  = (company.getElement('industries') != null ? company.getElement('industries').getElements() : []);
  var industries = '';//
  for(var i in industriesXml){
    var industryName = industriesXml[i].getElement('name').getText();
    industries += (industries == '' ? industryName : ', '+industryName);
  }
  
  var statusXml = company.getElement('status');
  var status = '';
  if(statusXml != null) status += statusXml.getElement('name').getText();//
  
  var blogRssUrlXml  = company.getElement('blog-rss-url');
  var blogRssUrl = '';
  if(blogRssUrlXml != null) blogRssUrl += blogRssUrlXml.getText();
  
  var twitterIdXml  = company.getElement('twitter-id');
  var twitterId = '';
  if(twitterIdXml != null)twitterId += twitterIdXml.getText();//
  
  var employeeCountRangeXml  =  company.getElement('employee-count-range');
  var employeeCountRange = '';
  if(employeeCountRangeXml != null)
    employeeCountRange += employeeCountRangeXml.getElement('name').getText();
  
  
  var specialties = '';
  var specialtiesXml  = (company.getElement('specialties') != null ? company.getElement('specialties').getElements() : []);
  for (var i in specialtiesXml) {
    specialty = specialtiesXml[i].getText();
    specialties += ', '+ specialty;
  }  
  specialties = specialties.substr(1);
  
  var locations = '';
  var locationsXml  = (company.getElement('locations') != null ? company.getElement('locations').getElements() : []);
  for(var i in locationsXml){
    var address = '';
    var addressXml = locationsXml[i].getElement('address');
    var addressparts = addressXml.getElements();
    for(var j in addressparts){
      address += (address == '' ? addressparts[j].getText() : ', '+addressparts[j].getText());
    }
    locations += (locations == '' ? address : '\n'+address);
  }
  
  var foundedYearXml  = company.getElement('founded-year');
  var foundedYear = '';
  if(foundedYearXml != null)foundedYear += foundedYearXml.getText();//
  
  return [id, name, numFollowers, companyType, industries, status, blogRssUrl, twitterId, employeeCountRange, specialties, locations, foundedYear];
}

//Sample xml
/*
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<companies total="1">
  <company>
    <id>167785</id>
    <name>Trulia</name>
  </company>
</companies>
OR
<company>
<id>1035</id>
<name>Microsoft</name>
</company>
*/

/*
 * Extract Company ID from XML
 */
function getCompanyId_(xmlRaw){
  var xml = Xml.parse(xmlRaw);
  var companies = xml.getElement();
  var total = companies.getAttribute('total');
  if (total != null) {
    var numCompanies = total.getValue();
    if(parseInt(numCompanies) > 0){
      var company = companies.getElements('company')[0];
      var id = company.getElement('id').getText();
      var name = company.getElement('name').getText();
      return {id:id, name:name};
    }  
  }
  else {
    // lets try other format
    var company = xml.getElement();
    if (company != null) {
      var id = company.getElement('id').getText();
      var name = company.getElement('name').getText();
      return {id:id, name:name};   
    }     
  }  
  return 'NOT FOUND';
}

/*
 * Get URL for Authorization (Step 1 of oAuth 2.0)
 */
function getURLForAuthorization(){
  STATE = generateRandomString(25); 
  Logger.log("STATE =" + STATE);
  UserProperties.setProperty(statePropertyName, STATE);
  return AUTHORIZE_URL + '?response_type=code&client_id='+CLIENT_ID+'&redirect_uri='+REDIRECT_URL +
    '&scope=r_fullprofile r_emailaddress r_network&state='+STATE;  
}

/*
 * Fetch the oAuth 2.0 Access Token and save as a script property
 */
function getAndStoreAccessToken(code){
  var parameters = {
     method : 'post',
     payload : 'client_id='+CLIENT_ID+'&client_secret='+CLIENT_SECRET+'&grant_type=authorization_code&redirect_uri='+REDIRECT_URL+'&code=' + code,
    muteHttpExceptions: false
   };
  
  var response = UrlFetchApp.fetch(TOKEN_URL,parameters).getContentText();   
  Logger.log(response);
  var tokenResponse = JSON.parse(response);
  
  //store the token for later retrival
  UserProperties.setProperty(tokenPropertyName, tokenResponse.access_token);
}

/*
 * Check if a valid oAuth 2.0 access token is available in script property
 */
function isTokenValid() {
  var token = UserProperties.getProperty(tokenPropertyName);
  Logger.log("Token Value = "+token);
  if(!token){ //if its empty or undefined
    return false;
  }
  return true; 
}

/*
 * Generate a random string
 */
function generateRandomString(n) {    
  var chars = ['a', 'b','c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
  chars.push('A', 'B', 'C', 'D', 'E', 'F','G','H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
  var randomString = '';  
  for (i=0; i < n; i++) {
    r  = Math.random();
    r = r * 61; 
    r = Math.round(r);  
    randomString = randomString + chars[r];
  }  
  return randomString;
}
