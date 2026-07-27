using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class JsonIso8601CultureDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Set the workbook's culture to invariant to ensure ISO formatting is not altered by locale
            workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

            // Prepare JSON data containing date strings in any format
            string jsonData = @"[
                { ""Id"": 1, ""Name"": ""Alice"", ""JoinDate"": ""2023-07-15T08:30:00"" },
                { ""Id"": 2, ""Name"": ""Bob"",   ""JoinDate"": ""2023-08-01T14:45:00"" }
            ]";

            // Configure JsonLayoutOptions to convert strings to dates and enforce ISO 8601 format
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true,
                // ISO 8601 format (date and time with 'T' separator)
                DateFormat = "yyyy-MM-ddTHH:mm:ss"
            };

            // Import JSON data into the first worksheet (lifecycle rule: load)
            JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

            // Prepare JsonSaveOptions to export the worksheet back to JSON
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                Indent = "    ", // optional pretty printing
                ExportArea = new CellArea { StartRow = 0, EndRow = 2, StartColumn = 0, EndColumn = 2 },
                HasHeaderRow = true
            };

            // Save the workbook as JSON (lifecycle rule: save)
            workbook.Save("ExportedIso8601.json", saveOptions);
        }
    }
}