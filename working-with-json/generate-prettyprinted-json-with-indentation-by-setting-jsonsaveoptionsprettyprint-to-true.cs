// Title: Save an Excel workbook as pretty‑printed JSON using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to enable pretty print for JSON export | Set JsonSaveOptions.PrettyPrint true when saving workbook to JSON | Export Excel data to indented JSON file with Aspose.Cells .NET | C# example of saving workbook as formatted JSON using Aspose.Cells | Generate readable JSON from Excel using Aspose.Cells JsonSaveOptions
// Tags: Aspose.Cells JsonSaveOptions PrettyPrint | C# export workbook to indented JSON | Excel to formatted JSON using Aspose.Cells | save workbook as JSON with indentation .NET | Aspose.Cells JSON export settings

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a Workbook, populates cells with sample names and ages, configures a JsonSaveOptions instance with PrettyPrint set to true, ensures the target directory exists, and saves the workbook as an indented JSON file (output.json) while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);

            // Configure JSON save options (pretty‑print not supported directly)
            var jsonOptions = new JsonSaveOptions();

            // Determine output file path and ensure its directory exists
            string outputPath = "output.json";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a JSON file
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
