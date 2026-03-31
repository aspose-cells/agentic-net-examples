using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main()
        {
            // Sample JSON representing Excel data (replace with actual JSON source)
            string jsonData = @"[
                { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
                { ""Name"": ""Alice"", ""Age"": 25, ""City"": ""London"" },
                { ""Name"": ""Bob"", ""Age"": 35, ""City"": ""Paris"" }
            ]";

            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Configure JSON layout options to treat the array as a table
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true   // Preserve column integrity
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonData, cells, 0, 0, layoutOptions);

            // Save the workbook as CSV (save rule)
            string outputCsvPath = "output.csv";
            workbook.Save(outputCsvPath, SaveFormat.Csv);

            Console.WriteLine($"JSON data has been converted to CSV and saved to '{outputCsvPath}'.");
        }
    }
}