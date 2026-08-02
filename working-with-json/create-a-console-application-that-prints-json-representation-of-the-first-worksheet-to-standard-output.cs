// Title: C# console app: Export the first worksheet to formatted JSON with Aspose.Cells
// Description: Creates a new Workbook, fills the first worksheet with sample data, defines a range, configures JsonSaveOptions (header row, empty cells as null, 4‑space indentation), and uses JsonUtility.ExportRangeToJson to print a pretty‑printed JSON string to the console inside a try‑catch block.
// Keywords: Aspose.Cells JSON export C# | JsonSaveOptions HasHeaderRow | ExportRangeToJson example | pretty print JSON from Excel .NET | console application Aspose.Cells JSON
// Common Searches: convert Aspose.Cells worksheet to JSON string C# | Aspose.Cells JsonSaveOptions indentation example | Export Excel range to JSON with header row .NET | C# console output JSON from first worksheet Aspose.Cells
// Developer Intent: Generate a JSON representation of the first worksheet and write it to standard output.
// Use Cases: Quickly expose worksheet data as JSON for web APIs or client‑side scripts. | Create readable JSON fixtures for unit tests or documentation. | Log Excel content in a structured format for debugging or audit trails.
// AI Prompts: Write a reusable method that takes a Worksheet and returns its JSON string using customizable JsonSaveOptions. | Modify the console program to accept a file path argument, load the workbook, and export the first sheet to JSON. | Extend the example to loop through all worksheets and save each as an indented JSON file.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExport
{
    // Creates a new Workbook, fills the first worksheet with sample data, defines a range, configures JsonSaveOptions (header row, empty cells as null, 4‑space indentation), and uses JsonUtility.ExportRangeToJson to print a pretty‑printed JSON string to the console inside a try‑catch block.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (including a header row)
                worksheet.Cells["A1"].PutValue("Name");
                worksheet.Cells["B1"].PutValue("Age");
                worksheet.Cells["A2"].PutValue("John");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["A3"].PutValue("Alice");
                worksheet.Cells["B3"].PutValue(25);

                // Define the range that covers the populated cells
                Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:B3");

                // Configure JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,          // First row contains column names
                    ExportEmptyCells = true,      // Include empty cells as null
                    Indent = "    "               // Pretty‑print with 4‑space indentation
                };

                // Export the range to a JSON string
                string json = JsonUtility.ExportRangeToJson(range, jsonOptions);

                // Output the JSON representation to the console
                Console.WriteLine(json);
            }
            catch (Exception ex)
            {
                // Log or display any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
