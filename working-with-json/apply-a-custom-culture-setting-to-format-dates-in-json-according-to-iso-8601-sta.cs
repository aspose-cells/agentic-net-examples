using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonCultureDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook culture to invariant to ensure culture‑independent formatting
            workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

            // Prepare JSON data containing date strings in a generic format
            string jsonData = @"[
                { ""Id"": 1, ""Name"": ""Alice"", ""JoinDate"": ""2023-05-15T08:30:00Z"" },
                { ""Id"": 2, ""Name"": ""Bob"",   ""JoinDate"": ""2023-06-20T14:45:00Z"" }
            ]";

            // Configure JSON import options to convert strings to dates and enforce ISO‑8601 format
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true,
                // ISO 8601 format (date and time in UTC)
                DateFormat = "yyyy-MM-ddTHH:mm:ssZ"
            };

            // Import the JSON data into the first worksheet starting at cell A1
            JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

            // Save the workbook to an Excel file (the dates will be stored using the specified format)
            workbook.Save("JsonWithIso8601Dates.xlsx");
        }
    }
}