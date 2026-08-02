using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonStringExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate cells with headers
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");

                // Populate cells with data (numeric values)
                cells["A2"].PutValue("Laptop");
                cells["B2"].PutValue(999.99);
                cells["A3"].PutValue("Phone");
                cells["B3"].PutValue(599.99);

                // Apply a custom number format to the price column (e.g., currency)
                Style priceStyle = workbook.CreateStyle();
                priceStyle.Number = 2; // Built‑in currency format
                cells["B2:B3"].SetStyle(priceStyle);

                // Define the range to export (including header row)
                AsposeRange exportRange = cells.CreateRange("A1:B3");

                // Configure export options to output all values as strings
                ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
                {
                    HasHeaderRow = true,      // Treat first row as header
                    ExportAsString = true,    // Force numeric values to be exported as strings
                    Indent = "    "           // Pretty‑print with 4‑space indentation
                };

                // Export the range to JSON using the configured options
                string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

                // Display the resulting JSON
                Console.WriteLine("Exported JSON with numeric values as strings:");
                Console.WriteLine(jsonOutput);
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File not found: {ex.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}