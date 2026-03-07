using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsMhtmlToJson
{
    public class Converter
    {
        public static void Run()
        {
            // Path to the source MHTML file that contains CSV data
            string mhtmlPath = "source.mhtml";

            // If the MHTML file does not exist, create a simple workbook and save it as MHTML
            if (!File.Exists(mhtmlPath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Name");
                ws.Cells["B1"].PutValue("Age");
                ws.Cells["A2"].PutValue("John");
                ws.Cells["B2"].PutValue(30);
                ws.Cells["A3"].PutValue("Jane");
                ws.Cells["B3"].PutValue(25);
                tempWb.Save(mhtmlPath, SaveFormat.MHtml);
            }

            // Load the MHTML file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.MHtml);
            Workbook workbook = new Workbook(mhtmlPath, loadOptions);

            // Access the first worksheet (assumed to contain the imported CSV)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the used range of the worksheet
            var usedRange = worksheet.Cells.MaxDisplayRange;

            // Configure JSON export options (export as Excel‑style JSON structure)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Write the JSON string to an output file
            File.WriteAllText("output.json", json);

            Console.WriteLine("MHTML CSV has been converted to JSON successfully.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            Converter.Run();
        }
    }
}