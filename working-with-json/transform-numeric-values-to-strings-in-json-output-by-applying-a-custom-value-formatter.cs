using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonNumericToStringDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with sample data containing numeric values
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(999.99);
            worksheet.Cells["A3"].PutValue("Phone");
            worksheet.Cells["B3"].PutValue(599.99);

            // Configure export options to output numeric values as strings
            ExportRangeToJsonOptions exportOptions = new ExportRangeToJsonOptions
            {
                HasHeaderRow = true,          // First row contains column names
                ExportAsString = true,        // Force all cell values to be exported as strings
                Indent = "    "               // Pretty‑print JSON with 4‑space indentation
            };

            // Define the range to export (including header)
            Aspose.Cells.Range exportRange = worksheet.Cells.CreateRange("A1:B3");

            // Export the range to a JSON string using the configured options
            string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, exportOptions);

            // Save the JSON string to a file
            File.WriteAllText("NumericAsString.json", jsonOutput);

            // Save the workbook to Excel for verification (optional)
            workbook.Save("NumericDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}