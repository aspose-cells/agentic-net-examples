using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExportDemo
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

                // Add header and a date value
                sheet.Cells["A1"].PutValue("Date");
                DateTime sampleDate = new DateTime(2023, 5, 15);
                sheet.Cells["A2"].PutValue(sampleDate);

                // Apply a custom date format to the cell (e.g., "dd-MM-yyyy")
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "dd-MM-yyyy";
                sheet.Cells["A2"].SetStyle(dateStyle);

                // Configure JSON save options
                JsonSaveOptions saveOptions = new JsonSaveOptions
                {
                    // Export cell values as strings so the custom format is preserved in JSON
                    ExportAsString = true,

                    // Optional: format the output JSON with indentation
                    Indent = "    ",

                    // Export the whole sheet (no specific ExportArea needed)
                    HasHeaderRow = true
                };

                // Export the worksheet range to JSON string
                Aspose.Cells.Range exportRange = sheet.Cells.CreateRange("A1:A2");
                string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, saveOptions);

                // Output the JSON to console
                Console.WriteLine("Exported JSON:");
                Console.WriteLine(jsonOutput);

                // Save the JSON to a file using the same options
                string jsonFilePath = "ExportedData.json";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(jsonFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(jsonFilePath, saveOptions);
                Console.WriteLine($"JSON file saved to: {jsonFilePath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}