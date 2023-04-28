

Step 1 Create or go to the google sheet that you would like to use. 

Step 2 Go to extensions and press on Apps Script 

![Screenshot 2023-04-27 at 7 26 17 PM](https://user-images.githubusercontent.com/81628855/235012361-c75d149e-f67c-47fa-875f-4daa8308bd9b.png)



Step 3 

Copy and paste this.....






























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
