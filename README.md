

XXXXXXXXXXX Step 1 Create or go to the google sheet that you would like to use. XXXXXXXXXXXX

XXXXXXXXXXX Step 2 Go to extensions and press on Apps Script XXXXXXXXXXX

![Screenshot 2023-04-27 at 7 26 17 PM](https://user-images.githubusercontent.com/81628855/235012361-c75d149e-f67c-47fa-875f-4daa8308bd9b.png)



XXXXXXXXXX Step 3 XXXXXXXXXX

Copy and paste this.....

var REFRESH_INTERVAL_MINUTES = 2;

function vinMakeInfo(vinNumber) { if ((vinNumber) && (vinNumber != "")) { try { vin = encodeURI(vinNumber) var url = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/"+vin+"?format=json" var response = UrlFetchApp.fetch(url); var w = JSON.parse(response.getContentText()); return w.Results[0].Make; } catch (e) {
} } return "" }

function vinModelInfo(vinNumber) { if ((vinNumber) && (vinNumber != "")) { try { vin = encodeURI(vinNumber) var url = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/"+vin+"?format=json" var response = UrlFetchApp.fetch(url); var w = JSON.parse(response.getContentText()); return w.Results[0].Model; } catch (e) {
} } return "" }

function vinModelYearInfo(vinNumber) { if ((vinNumber) && (vinNumber != "")) { try { vin = encodeURI(vinNumber) var url = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/"+vin+"?format=json" var response = UrlFetchApp.fetch(url); var w = JSON.parse(response.getContentText()); return w.Results[0].ModelYear; } catch (e) {
} } return "" }

function vinAuctionInfo(vinNumber) { if ((vinNumber) && (vinNumber != "")) { try { vin = encodeURI(vinNumber) var url = "https://pinpoint-us-api.cox2m.com/v1/association?siteId=lmaa&size=1000&q=" + vin var response = UrlFetchApp.fetch(url); var w = JSON.parse(response.getContentText()); var total = w.total

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
} return "" }

function refresh() { SpreadsheetApp.getUi().alert("Refreshing auction states"); var ss = SpreadsheetApp.getActiveSpreadsheet(); }

function onOpen() { try { SpreadsheetApp.getUi().alert("Testing periodic update"); ScriptApp.newTrigger('refresh') .timeBased() .everyMinutes(REFRESH_INTERVAL_MINUTES) .create(); SpreadsheetApp.getUi().alert("Addred refresh trigger for ", REFRESH_INTERVAL_MINUTES.toString() + " Minutes"); } catch (e) { SpreadsheetApp.getUi().alert(e.toString()); } }




It should look like this......


![Screenshot 2023-04-27 at 8 08 53 PM](https://user-images.githubusercontent.com/81628855/235015882-279c0667-c212-468b-9570-a714d57b7649.png)







XXXXXXXXXX Step 4 XXXXXXXXXX
Save project 
Then press run. 
Grant access.

![Screenshot 2023-04-27 at 8 19 11 PM](https://user-images.githubusercontent.com/81628855/235023526-744796ef-bbad-41a9-b356-e876fc60cea1.png)




XXXXXXXXXX Step 5 XXXXXXXXXX

Copy and paste these functions.

Year 
=vinModelYearInfo(insert here the cell where you are going to place the VIN )


Make 

=vinMakeInfo(insert here the cell where you are going to place the VIN)


Model 
=vinModelInfo(=vinModelInfo(insert here the cell where you are going to place the VIN)



Enjoy!!!!! 



If you are still struggling just message me and I will help you out...........I got you

























# Autopopulate-Year-Make-and-Model-Google-Sheets-
These functions will call the NHTSA API and Auto-populate the cells on Google sheets


If you work with automobiles, you know that the VIN number is a crucial piece of information. It tells you everything from the make and model of the car to its year of manufacture. Coping and pasting year, make and model can get quite tedious. With the NHTSA API and Google Sheets, you can automate this process and quickly get the information you need.
The NHTSA API is a web service provided by the National Highway Traffic Safety Administration. It allows you to query the NHTSA database and retrieve information about vehicles, including their make, model, and year. By using this API in conjunction with Google Sheets, you can create a powerful tool for tracking information about automobiles.
In this article, we'll walk through how to use the NHTSA API to automatically populate cells in Google Sheets with information about a vehicle's make, model, and year. We'll use three functions: vinMakeInfo, vinModelInfo, and vinModelYearInfo.
vinMakeInfo(vinNumber) is the first function. It takes a VIN number as its input and returns the make of the vehicle. Here's how it works:
1.	Check that the VIN number is not empty.
2.	Encode the VIN number using encodeURI().
3.	Construct a URL to query the NHTSA API, passing in the encoded VIN number and setting the format to JSON.
4.	Use the UrlFetchApp.fetch() method to make the API call and retrieve the response.
5.	Parse the response as JSON using JSON.parse().
6.	Return the make of the vehicle from the response.
The second function, vinModelInfo(vinNumber), is similar to vinMakeInfo but returns the model of the vehicle instead. The third function, vinModelYearInfo(vinNumber), returns the year of manufacture.
With these three functions, you can quickly and easily retrieve information about a vehicle using its VIN number. To use them in Google Sheets, simply enter the VIN number in one cell and then use the functions in other cells to retrieve the make, model, and year.
Automating this process can save you time and reduce errors. Instead of manually looking up information about a vehicle, you can rely on the NHTSA API to provide accurate and up-to-date information. And with Google Sheets, you can easily share this information with others and collaborate on tracking data about automobiles.
In conclusion, using the NHTSA API and Google Sheets together can be a powerful tool for anyone who works with automobiles. By automating the process of retrieving information about vehicles, you can save time and reduce errors. So why not give it a try and see how it can benefit you?
