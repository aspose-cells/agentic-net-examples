// Title: Convert JSON to CSV with a semicolon delimiter using Aspose.Cells in C#
// Common Searches: C# set semicolon as CSV separator when exporting JSON with Aspose.Cells | How to use TxtSaveOptions to define a custom delimiter for CSV in Aspose.Cells | Aspose.Cells JsonLayoutOptions import JSON and save as CSV with semicolon | Convert JSON array to CSV file using a semicolon delimiter in .NET | Export worksheet to CSV with ';' separator using Aspose.Cells C#
// Tags: Aspose.Cells JsonLayoutOptions CSV semicolon | TxtSaveOptions custom separator C# | Import JSON into worksheet Aspose.Cells | Export worksheet to CSV with custom delimiter | C# JSON to CSV conversion Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

namespace JsonToCsvWithSemicolon
{
    // The example reads a JSON file, configures JsonLayoutOptions (including array handling and type conversion), imports the data into a worksheet, sets TxtSaveOptions.Separator to ';' with UTF‑8 encoding, and saves the workbook as a CSV file using the semicolon delimiter.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: JsonToCsvWithSemicolon <inputJsonFilePath> <outputCsvFilePath>");
                return;
            }

            string jsonFilePath = args[0];
            string csvFilePath = args[1];

            if (!File.Exists(jsonFilePath))
            {
                Console.WriteLine($"Error: JSON file not found at '{jsonFilePath}'.");
                return;
            }

            try
            {
                // Read JSON content
                string jsonContent = File.ReadAllText(jsonFilePath);

                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure JsonLayoutOptions
                JsonLayoutOptions jsonOptions = new JsonLayoutOptions
                {
                    // Process arrays as tables so the output resembles a CSV structure
                    ArrayAsTable = true,
                    // Ignore titles to produce a pure data CSV
                    IgnoreTitle = true,
                    // Convert numeric/date strings to proper types
                    ConvertNumericOrDate = true,
                    // Define date and number formats if needed
                    DateFormat = "yyyy-MM-dd",
                    NumberFormat = "#,##0.00"
                };

                // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
                JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, jsonOptions);

                // Prepare TXT save options (used for CSV) with semicolon as the delimiter
                TxtSaveOptions txtOptions = new TxtSaveOptions
                {
                    Separator = ';',
                    Encoding = Encoding.UTF8
                    // SaveFormat is read‑only and defaults to Csv; no need to set it.
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(csvFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as CSV using the defined options
                workbook.Save(csvFilePath, txtOptions);

                Console.WriteLine("Successfully converted JSON to CSV with semicolon delimiter.");
                Console.WriteLine($"Input JSON:  {jsonFilePath}");
                Console.WriteLine($"Output CSV:  {csvFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
