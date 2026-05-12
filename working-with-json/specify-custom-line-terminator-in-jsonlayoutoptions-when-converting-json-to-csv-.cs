using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Sample JSON data
        string json = @"{
            ""Products"": [
                { ""ID"": 1, ""Name"": ""Laptop"", ""Price"": 999.99 },
                { ""ID"": 2, ""Name"": ""Phone"",  ""Price"": 599.99 }
            ]
        }";

        // Configure JSON layout options for import
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,          // Treat arrays as tables
            ConvertNumericOrDate = true   // Convert numbers and dates automatically
        };

        // Create a new workbook and import the JSON data into the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        JsonUtility.ImportData(json, worksheet.Cells, 0, 0, layoutOptions);

        // Define the CSV file path
        string csvPath = "Products.csv";

        // Save the workbook as CSV (default line separator may be environment dependent)
        workbook.Save(csvPath, SaveFormat.Csv);

        // Ensure Windows‑style line endings ("\r\n") for compatibility
        string csvContent = File.ReadAllText(csvPath);
        // Replace any lone LF with CRLF
        csvContent = csvContent.Replace("\r\n", "\n")   // normalize existing CRLF to LF
                               .Replace("\n", "\r\n"); // convert all LF to CRLF
        File.WriteAllText(csvPath, csvContent);
    }
}