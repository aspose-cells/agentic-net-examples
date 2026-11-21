// Title: How to use a custom value formatter in Aspose.Cells for .NET to export numeric Excel cells as strings in JSON
// Common Searches: how to export Excel numeric cells as strings in JSON using Aspose.Cells C# | Aspose.Cells JsonSaveOptions custom value formatter example for numeric to string conversion | C# convert workbook numeric values to string when saving as JSON with Aspose.Cells
// Tags: Aspose.Cells JSON custom value formatter | numeric cells to string in JSON export | C# Excel to JSON numeric string conversion | JsonSaveOptions custom formatting for numbers | export workbook values as strings in JSON

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

// Demonstrates setting a custom value formatter on JsonSaveOptions so that numeric cell values are written as strings in the generated JSON file when saving a workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);      // numeric value
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92.5);    // numeric value

            // Configure JSON save options (no custom formatter needed)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Determine output path and ensure the directory exists
            string outputPath = "output.json";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as JSON
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
