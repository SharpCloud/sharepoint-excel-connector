## sharepoint-excel-connector
Uses the Excel REST API in SharePoint Online to read an Excel spreadsheet and save the data to a SharpCloud story.

#### App.config file settings

**sharpCloudStoryID**
The ID of the SharpCloud story that will be updated with data from the spreadsheet

**SharpCloudUsername**
**SharpCloudPassword**
The username and password of a user that has edit rights on the SharpCloud story

**sharePointUsername**
**sharePointPassword**
The username and password of the Microsoft Account with rights to read the Excel spreadhsheet in SharePoint

**sharePointSite**
The domain where the SharePoint site is hosted, eg. https://sharpcloud.sharepoint.com

**sharePointDirectory**
The SharePoint directory where the Excel spreadsheet is stored

**spreadSheetName**
The name of the Excel spreadsheet.

**spreadSheetTableName**
The name of the table within the Excel spreadsheet where the data we will process is stored.

#### Resources

For more information on using the Excel REST API in SharePoint Online see this article:
http://blogs.office.com/2013/12/17/excel-rest-api-in-sharepoint-online/

And for more information on the format of the Resources URI for Excel Services REST API:
https://msdn.microsoft.com/en-us/library/office/ff394530(v=office.14).aspx
