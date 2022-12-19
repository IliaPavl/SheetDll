using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace Sheet
{
    [ComVisible(true)]
    public interface ISheetHelper
    {
        void SetProperty(string sheetId, string sheetName, string passJsonKey, string nameProgect);
        void PrintEntries(string[,] values);
        void DeleteEntry(string start, string end);
        string ReadEntry(string point);
        string[,] ReadEntries(string start, string end);
        void CreateEntry(string start, string end, string[] listValues);
        void UpdateEntry(string point, string[] listValues);
        void PrintNotNullEntries(string[,] values);
        string[,] ReadCommand(string programString);
        void UpdateCommand(string command, string[] listValues);
        void CreateCommand(string command, string[] listValues);
        void DeleteCommand(string command);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class SheetHelper : ISheetHelper
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName;
        static string SpreadsheetId;
        static string sheet;
        static SheetsService service;

        //Установка настроек ______________________________________________________________________
        public void SetProperty(string sheetId, string sheetName, string passJsonKey, string nameProgect)
        {
            try
            {
                ApplicationName = nameProgect;
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
            catch (Exception e) { Console.WriteLine("Error set settings:" + e); }
        }

        //Красивый вывод __________________________________________________________________________
        public void PrintEntries(string[,] values)
        {
            if (values != null && values.Length > 0)
            {
                for (int j = 0; j < values.GetLength(0); j++)
                    for (int i = 0; i < values.GetLength(1); i++)
                        if (i == values.GetLength(1) - 1)
                        {
                            Console.Write("{0}", values[j, i]);
                            Console.WriteLine("");
                        }
                        else
                            Console.Write("{0} | ", values[j, i]);
                Console.WriteLine("");
            }
            else
                Console.WriteLine("No data found.");
        }
        
        public void PrintNotNullEntries(string[,] values)
        {
            if (values != null && values.Length > 0)
            {
                for (int j = 0; j < values.GetLength(0); j++)
                {
                    for (int i = 0; i < values.GetLength(1); i++)
                        if (i == values.GetLength(1) - 1)
                            if (values[j, i] != null)
                                Console.Write("{0}", values[j, i]);
                            else if (values[j, i] != null)
                                Console.Write("{0} | ", values[j, i]);
                    Console.WriteLine("");
                }
                Console.WriteLine("");
            }
            else
                Console.WriteLine("No data found.");
        }

        //Круд операции ___________________________________________________________________________
        public string[,] ReadEntries(string start, string end)
        {
            string range = $"{sheet}!{start}:{end}";
            return ReadCommand(range);
        }

        public string ReadEntry(string point)
        {
            var range = $"{sheet}!{point}:{point}";
            string[,] list = ReadCommand(range);
            if (list != null && list.GetLength(0) > 0 && list.GetLength(1) > 0)
                return list[0, 0];
            else
                return null;
        }

        public void DeleteEntry(string start, string end)
        {
            var range = $"{sheet}!{start}:{end}";
            DeleteCommand(range);
        }

        public void CreateEntry(string start, string end, string[] listValues)
        {
            var range = $"{sheet}!{start}:{end}";
            CreateCommand(range, listValues);
        }

        public void UpdateEntry(string point, string[] listValues)
        {
            var range = $"{sheet}!{point}:{point}";
            UpdateCommand(range, listValues);
        }

        //Любой запрос сюда подставляеш ___________________________________________________________
        public string[,] ReadCommand(string programString)
        {
            try
            {
                SpreadsheetsResource.ValuesResource.GetRequest request =
                         service.Spreadsheets.Values.Get(SpreadsheetId, programString);
                IList<IList<object>> obj = request.Execute().Values;
                string[,] list = null;
                int firstColumn = obj.Count, endColumn = -1;
                if (obj != null && obj.Count > 0)
                {
                    for (int j = 0; j < obj.Count; j++)
                        if (endColumn < obj[j].Count)
                            endColumn = obj[j].Count;

                    list = new string[firstColumn, endColumn];

                    for (int j = 0; j < obj.Count; j++)
                        for (int i = 0; i < obj[j].Count; i++)
                            list[j, i] = (string)obj[j][i];
                }
                return list;
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); return null; }
        }
       
        public void UpdateCommand(string command, string[] listValues)
        {
            try
            {
                var valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { listValues };
                var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, command);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = updateRequest.Execute();
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); }
        }

        public void CreateCommand(string command, string[] listValues)
        {
            try
            {
                var valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { listValues };
                var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, command);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = updateRequest.Execute();
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); }
        }

        public void DeleteCommand(string command)
        {
            try
            {
                var requestBody = new ClearValuesRequest();
                var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, command);
                var deleteReponse = deleteRequest.Execute();
            }
            catch (Exception e)
            { Console.WriteLine("Error: " + e); }
        }
    }
}
