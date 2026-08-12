// Title: C# – Convert JSON to Semicolon‑Delimited CSV with Aspose.Cells
// Description: Demonstrates how to import JSON into a worksheet using JsonLayoutOptions (ArrayAsTable, ConvertNumericOrDate) and export it as a CSV file with a semicolon separator via TxtSaveOptions. Includes directory creation and basic error handling.
// Keywords: Aspose.Cells JSON to CSV | C# semicolon CSV delimiter | JsonLayoutOptions ArrayAsTable | TxtSaveOptions custom separator | .NET export JSON as CSV | European locale CSV format
// Common Searches: Aspose.Cells set CSV delimiter to semicolon | C# convert JSON array to CSV with custom separator | JsonUtility import JSON and save as CSV Aspose | How to use TxtSaveOptions Separator property | Semicolon‑delimited CSV export .NET
// Developer Intent: Export JSON data to a CSV file using a semicolon as the field separator with Aspose.Cells.
// Use Cases: Create CSV reports for European systems that require ';' as the delimiter. | Transform JSON APIs into semicolon‑delimited files for legacy import pipelines. | Automate data exchange where numeric strings must be converted to numbers before CSV export.
// AI Prompts: Generate C# code that reads JSON, loads it into an Aspose.Cells workbook, and saves it as a semicolon‑delimited CSV. | Explain the impact of JsonLayoutOptions.ArrayAsTable and ConvertNumericOrDate on the worksheet before CSV conversion. | Provide step‑by‑step instructions to configure TxtSaveOptions for a custom CSV separator in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to import JSON into a worksheet using JsonLayoutOptions (ArrayAsTable, ConvertNumericOrDate) and export it as a CSV file with a semicolon separator via TxtSaveOptions. Includes directory creation and basic error handling.
class JsonToCsvWithSemicolon
{
    static void Main()
    {
        try
        {
            // Sample JSON data
            string json = @"{
                ""People"": [
                    { ""Name"": ""John"", ""Age"": 30 },
                    { ""Name"": ""Alice"", ""Age"": 25 }
                ]
            }";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure JSON layout options (optional settings)
            JsonLayoutOptions jsonOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,          // Treat arrays as tables
                ConvertNumericOrDate = true   // Convert numeric/date strings
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(json, worksheet.Cells, 0, 0, jsonOptions);

            // Set CSV save options with semicolon as the delimiter
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ';'               // Use semicolon as CSV delimiter
                // ConvertNumericData property is not available; numeric conversion is handled during import
            };

            // Define output file path
            string outputPath = "output.csv";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the worksheet as a CSV file using the defined options
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"CSV file saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
