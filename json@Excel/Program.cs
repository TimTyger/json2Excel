using Aspose.Cells;
using Aspose.Cells.Utility;
using GemBox.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace json_Excel
{
    class Program
    {
        static void Main(string JsonFileName, string OutputFileName)
        {
            
            if (!JsonFileName.ToLower().Contains(".json"))
            {
                JsonFileName = $"{JsonFileName}.json";
            }
            if (!OutputFileName.ToLower().Contains(".xlsx"))
            {
                OutputFileName = $"{JsonFileName}.xlsx";
            }
            converter(JsonFileName, OutputFileName);
         
        }

        static void converter(string input, string output)
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            // Read JSON File
            string jsonInput = File.ReadAllText("Userdetails.json");

        // Set JsonLayoutOptions
        JsonLayoutOptions options = new JsonLayoutOptions();
        options.ArrayAsTable = true;

            // Import JSON Data
            JsonUtility.ImportData(jsonInput, worksheet.Cells, 0, 0, options);

            // Save Excel file
            workbook.Save("Import-Data-JSON-To-Excel.xlsx");
        }
    }
}
