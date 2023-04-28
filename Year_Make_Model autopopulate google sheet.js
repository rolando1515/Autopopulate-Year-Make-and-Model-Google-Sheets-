// These functions will call the NHTSA API and Auto-populate the cells on Google sheets
// To run it, take out the comments, copy and paste the code on App scripps.
// Then manually insert the function.
// Example   =vinModelInfo(F2) 
//           =vinMakeInfo(F2)
//           =vinModelYearInfo(F2)   ------> f2 is the cell where you insert the VIN 
// Drag and drop until desired cell



var REFRESH_INTERVAL_MINUTES = 2;


function vinMakeInfo(vinNumber)
{
  if ((vinNumber) && (vinNumber != ""))
  {
    try
    {
      vin = encodeURI(vinNumber)
      var url = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/"+vin+"?format=json"
      var response = UrlFetchApp.fetch(url);
      var w = JSON.parse(response.getContentText());
      return w.Results[0].Make;
    }
    catch (e)
    {      
    }
  }
  return ""
}


function vinModelInfo(vinNumber)
{
  if ((vinNumber) && (vinNumber != ""))
  {
    try
    {
      vin = encodeURI(vinNumber)
      var url = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/"+vin+"?format=json"
      var response = UrlFetchApp.fetch(url);
      var w = JSON.parse(response.getContentText());
      return w.Results[0].Model;
    }
    catch (e)
    {      
    }
  }
  return ""
}


function vinModelYearInfo(vinNumber)
{
  if ((vinNumber) && (vinNumber != ""))
  {
    try
    {
      vin = encodeURI(vinNumber)
      var url = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/"+vin+"?format=json"
      var response = UrlFetchApp.fetch(url);
      var w = JSON.parse(response.getContentText());
      return w.Results[0].ModelYear;
    }
    catch (e)
    {      
    }
  }
  return ""
}


function vinAuctionInfo(vinNumber)
{
  if ((vinNumber) && (vinNumber != ""))
  {
    try
    {
      vin = encodeURI(vinNumber)
      var url = "https://pinpoint-us-api.cox2m.com/v1/association?siteId=lmaa&size=1000&q=" + vin
      var response = UrlFetchApp.fetch(url);
      var w = JSON.parse(response.getContentText());
      var total = w.total
      
      try
      {
        total = parseInt(total);
      }
      catch (e)
      {
        return "Invalid count in API result";
      }
      
      if(total != 0)
      {
        return 'At Auction';
      }
      else
      {
        return 'Waiting Arrival';
      }
    }
    catch (e)
    {
      return "Look VIN Manually";
    }
  }
  return ""
}


function refresh() {
  SpreadsheetApp.getUi().alert("Refreshing auction states");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
}


function onOpen() {
  try {
    SpreadsheetApp.getUi().alert("Testing periodic update");
    ScriptApp.newTrigger('refresh')
        .timeBased()
        .everyMinutes(REFRESH_INTERVAL_MINUTES)
        .create();
    SpreadsheetApp.getUi().alert("Addred refresh trigger for ", REFRESH_INTERVAL_MINUTES.toString() + " Minutes");
  }
  catch (e)
  {
    SpreadsheetApp.getUi().alert(e.toString());
  }
}
