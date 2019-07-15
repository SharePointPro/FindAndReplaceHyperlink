# FindAndReplaceHyperlink
After being contracted for a SharePoint On premesis to SharePoint Online migration, it was discovered that many word docx hyperlinks had been broken. 

This app was created to update hyperlinks in docxs in bulk.

*USAGE*

Create a targets.txt fiole in the application directy, the targets.txt file contains a comma seperated list of target,destination

Exmaple targets.txt file
https://www.bing.com.au,https://www.google.com.au
https://www.myspace.com,https://www.facebook.com

The above targets.txt file will replace all hyperlinks pointing to bing with google, and all hyperlinks pointing to myspace with facebook.

It will keep the relative urls. for example if a hyperlink is https://www.bing.com.au/info/file.aspx it will replace this with https://www.google.com.au/file.aspx

*Aspose license is required*. 
The application is dependant on aspose word, and will not work without a paid license, one can be purchased from here https://www.aspose.com/

