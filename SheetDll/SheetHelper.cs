using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace SheetDll
{
    public class SheetHelper
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Current1132 Legislators";
        static string SpreadsheetId;
        static string sheet;
        static SheetsService service;

        public SheetHelper()
        {
        }
        public SheetHelper(string sheetId, string sheetName,string passJsonKey)
        {
            setProperty( sheetId,  sheetName,  passJsonKey);
        }

        public void setProperty(string sheetId, string sheetName, string passJsonKey) {
            SpreadsheetId = sheetId;
            sheet = sheetName;
            GoogleCredential credential;
            using (var stream = new FileStream(passJsonKey, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        
        public void Print(IList<IList<object>> values)
        {
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                    for (int i = 0; i < row.Count; i++)
                        if (i == row.Count - 1)
                        {
                            Console.Write("{0}", row[i]);
                            Console.WriteLine("");
                        }
                        else
                            Console.Write("{0} | ", row[i]);
                Console.WriteLine("");
            }
            else
                Console.WriteLine("No data found.");
        }

        public IList<IList<object>> ReadRange(string start, string end)
        {
            string range = $"{sheet}!{start}:{end}";
            return getResponse(range);
        }

        public string ReadOne(string point)
        {
            var range = $"{sheet}!{point}:{point}";
            IList<IList<object>> list = getResponse(range);
            string value = "";
            if (list != null && list.Count > 0)
            {
                foreach (var row in list)
                {
                    for (int i = 0; i < row.Count; i++) { value = (string)row[i]; break; }
                    break;
                }

            }
            else
                return "empty";
            return value;
        }

        public  IList<IList<object>> getResponse(string programString)
        {
            try
            {
                SpreadsheetsResource.ValuesResource.GetRequest request =
                         service.Spreadsheets.Values.Get(SpreadsheetId, programString);

                return request.Execute().Values;
            }
            catch (Exception e) { Console.WriteLine("Error: " + e); return null; }

        }

        public void DeleteEntry(string start, string end)
        {
            try
            {
                var range = $"{sheet}!{start}:{end}";
                var requestBody = new ClearValuesRequest();
                var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
                var deleteReponse = deleteRequest.Execute();
            }
            catch (Exception e) { Console.WriteLine("Error: " + e); }
        }

        //format listValues:  new List<object>() { "Hello!", "This", "was", "insertd", "via", "C#" };
        public void CreateEntry(string start, string end, List<object> listValues)
        {
            try
            {
                var range = $"{sheet}!{start}:{end}";
                var valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { listValues };
                var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = appendRequest.Execute();
            }
            catch (Exception e) { Console.WriteLine("Error: " + e); }

        }

        //format listValues:  new List<object>() { "updated" };
        public void UpdateEntry(string point, List<object> listValues)
        {
            try
            {
                var range = $"{sheet}!{point}:{point}";
                var valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { listValues };
                var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = updateRequest.Execute();
            }
            catch (Exception e) { Console.WriteLine("Error: " + e); }
        }
    }
}