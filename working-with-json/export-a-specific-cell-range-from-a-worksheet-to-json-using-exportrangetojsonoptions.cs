// Title: Export a selected cell range from an Aspose.Cells worksheet to an indented JSON file using C#
// AI Prompts: Generate C# code that uses Aspose.Cells to export the range A1:C3 from a worksheet into a DataTable, then serialize it with System.Text.Json into a pretty‑printed JSON file. | Show how to ensure the target directory exists and write the formatted JSON string to ExportedRange.json after exporting a worksheet range with Aspose.Cells.
// Common Searches: aspocells c# export worksheet range A1:C3 to json file with indentation | how to use ExportDataTable and System.Text.Json to convert Excel range to JSON in C# | save selected cells from Aspose.Cells workbook as pretty printed JSON using .NET | C# example for exporting Excel range to JSON with Aspose.Cells and System.Text.Json
// Tags: aspocells exportrange to json c# | exportdata table json serialization aspocells | write indented json file from excel range c# | system.text.json datatable serialization aspocells | create output directory before writing json c#

using System;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example creates a workbook, fills cells A1:C3 with sample data, exports that range to a DataTable using Aspose.Cells, serializes the DataTable to indented JSON with System.Text.Json, ensures the output folder exists, and writes the result to ExportedRange.json.
class ExportRangeToJsonExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("City");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["C2"].PutValue("New York");
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["C3"].PutValue("London");

            // Define the cell range to export (A1:C3)
            // Export the range to a DataTable first
            DataTable dt = sheet.Cells.ExportDataTable(0, 0, 3, 3, true);

            // Serialize the DataTable to JSON
            string json = JsonSerializer.Serialize(dt, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON string to a file
            string outputPath = "ExportedRange.json";

            // Ensure the directory exists (if a directory part is present)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(outputPath, json);

            Console.WriteLine("Range exported to JSON successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
