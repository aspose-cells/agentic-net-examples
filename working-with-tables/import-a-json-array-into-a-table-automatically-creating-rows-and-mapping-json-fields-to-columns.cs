// Title: C# – Import JSON Array into Excel as a Table with Aspose.Cells
// Description: This example shows how to use Aspose.Cells for .NET to read a JSON array, automatically create column headers from the object fields, and insert each element as a row in an Excel worksheet. By enabling JsonLayoutOptions.ArrayAsTable and calling JsonUtility.ImportData, the code builds a formatted table starting at cell A1 and saves it as an .xlsx file.
// Keywords: Aspose.Cells | C# JSON to Excel | JsonUtility ImportData | ArrayAsTable | Excel table from JSON | convert JSON array | Aspose.Cells .NET | JSON import C# | Excel automation | US Canada UK sample data
// Common Searches: Aspose.Cells import JSON array as table | JsonUtility ImportData C# example | Convert JSON to Excel table using .NET | Automatic column creation from JSON with Aspose.Cells | Set JsonLayoutOptions.ArrayAsTable
// Developer Intent: Create an Excel worksheet where each object in a JSON array becomes a row and each property becomes a column automatically.
// Use Cases: Turn API‑returned JSON data into an Excel table for quick analysis and reporting. | Migrate legacy JSON files to spreadsheets without manually defining column mappings. | Build a reusable utility that imports dynamic JSON structures into Excel for downstream business processes.
// AI Prompts: Write C# code that uses Aspose.Cells to import a JSON array into a worksheet as a table with automatic column creation. | Explain how JsonLayoutOptions.ArrayAsTable works and how to change the starting cell for the import. | Suggest error‑handling and performance strategies for importing large JSON payloads with JsonUtility.ImportData.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonImportExample
{
    // This example shows how to use Aspose.Cells for .NET to read a JSON array, automatically create column headers from the object fields, and insert each element as a row in an Excel worksheet. By enabling JsonLayoutOptions.ArrayAsTable and calling JsonUtility.ImportData, the code builds a formatted table starting at cell A1 and saves it as an .xlsx file.
    public class JsonToTableImporter
    {
        public static void Run()
        {
            try
            {
                // Sample JSON array – each object will become a row in the worksheet
                string json = @"[
                    { ""Name"": ""Alice"", ""Age"": 30, ""Country"": ""USA"" },
                    { ""Name"": ""Bob"",   ""Age"": 25, ""Country"": ""Canada"" },
                    { ""Name"": ""Carol"", ""Age"": 28, ""Country"": ""UK"" }
                ]";

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure layout options to treat the JSON array as a table
                JsonLayoutOptions options = new JsonLayoutOptions
                {
                    ArrayAsTable = true   // Enables automatic column creation and row insertion
                };

                // Import the JSON data starting at cell A1 (row 0, column 0)
                JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);

                // Define output file path
                string outputPath = "JsonArrayTable.xlsx";

                // Save the workbook to an Excel file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during JSON import or workbook save: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                JsonToTableImporter.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
