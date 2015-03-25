using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Text;
using System.Security;
using System.Web.Script.Serialization;
using Microsoft.SharePoint.Client;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;

namespace SharePointExcelConnector
{
    //Define data contracts for the JSON that will be returned from Excel Services
    [DataContract]
    public class CellValue
    {
        [DataMember]
        public object v { get; set; }
        [DataMember]
        public object fv { get; set; }
    }

    [DataContract]
    public class TableValue
    {
        [DataMember]
        public string name { get; set; }
        [DataMember]
        public CellValue[][] rows { get; set; }
    }

    class Program
    {

        private static void Main(string[] args)
        {
            UpdateStory();

            //Code for initiating from a URL when in an Azure Web Job
            /*
            JobHost h = new JobHost();
            h.Call(typeof (Program).GetMethod("UpdateStory"));
             * */
        }

        public static void UpdateStory()
        {

            string sharpCloudStoryID = ConfigurationManager.AppSettings["sharpCloudStoryID"];
            string sharpCloudUsername = ConfigurationManager.AppSettings["sharpCloudUsername"];
            string sharpCloudPassword = ConfigurationManager.AppSettings["sharpCloudPassword"];
            string spreadSheetName = ConfigurationManager.AppSettings["spreadSheetName"];
            string spreadSheetTableName = ConfigurationManager.AppSettings["spreadSheetTableName"];
            string sharePointUsername = ConfigurationManager.AppSettings["sharePointUsername"];
            string sharePointPassword = ConfigurationManager.AppSettings["sharePointPassword"];
            string sharePointSite = ConfigurationManager.AppSettings["sharePointSite"];
            string sharePointDirectory = ConfigurationManager.AppSettings["sharePointDirectory"];

            var securePassword = new SecureString();
            foreach (var character in sharePointPassword)
            {
                securePassword.AppendChar(character);
            }
            securePassword.MakeReadOnly();

            //create the SharePointOnlineCredentials object for authenticating with SharePoint
            var sharePointCredentials = new SharePointOnlineCredentials(sharePointUsername, securePassword);

            //construct the URI to access the Excel Services REST API
            //see https://msdn.microsoft.com/en-us/library/office/ff394530(v=office.14).aspx for more details
            string url = string.Format("https://{0}/_vti_bin/ExcelRest.aspx/personal/{1}/Documents/{2}/Model/Tables('{3}')?$format=json", sharePointSite, sharePointDirectory, spreadSheetName, spreadSheetTableName);

            var webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Credentials = sharePointCredentials;
            //The following line is necessary in SharePoint Online for every API request
            //it forces SharePoint online to challenge the client for credentials
            webRequest.Headers["X-FORMS_BASED_AUTH_ACCEPTED"] = "f";
            var webResponse = (HttpWebResponse)webRequest.GetResponse();

            String responseString;

            using (Stream stream = webResponse.GetResponseStream())
            {
                var reader = new StreamReader(stream, Encoding.UTF8);
                responseString = reader.ReadToEnd();
            }

            if (string.IsNullOrEmpty(responseString) == false)
            {

                //serialize the JSON into a TableValue type that we can use in code
                var serializer = new JavaScriptSerializer();
                var results = serializer.Deserialize<TableValue>(responseString);

                int rowCount = results.rows.Length;
                int colCount = results.rows[0].Length;

                //arrayValues is the array we construct and pass to the SharpCloud API to create/update the story
                var arrayValues = new string[rowCount, colCount];

                //look at each cell in the response to populate the arrayValues
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    var excelRow = results.rows[rowIndex];
                    for (int columnIndex = 0; columnIndex < colCount; columnIndex++)
                    {
                        if (excelRow[columnIndex] != null)
                        {
                            if (excelRow[columnIndex].v != null || excelRow[columnIndex].fv != null)
                            {
                                arrayValues[rowIndex, columnIndex] = excelRow[columnIndex].v.ToString();
                            }
                        }
                    }
                }

                var _client = new SharpCloudApi(sharpCloudUsername, sharpCloudPassword);
                Story rm = _client.LoadStory(sharpCloudStoryID);

                //UpdateStoryWithArray is where all the magic happens, this method takes an array and generates items/attributes etc in the SharpCloud story
                string errorMessage;
                if (rm.UpdateStoryWithArray(arrayValues, true, out errorMessage))
                {
                    rm.Save();
                }
            }

            Environment.Exit(0);
            
        }

    }
}
