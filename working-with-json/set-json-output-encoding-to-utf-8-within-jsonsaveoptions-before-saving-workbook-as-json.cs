// Title: Set UTF-8 encoding on JsonSaveOptions before saving an Aspose.Cells workbook as JSON in C#
// AI Prompts: Write C# code that assigns Encoding.UTF8 to JsonSaveOptions and saves a workbook to a JSON file using Aspose.Cells. | Show how to adapt an Aspose.Cells example so the generated JSON file is UTF‑8 encoded.
// Common Searches: Aspose.Cells how to export workbook to JSON with UTF-8 encoding in .NET | C# JsonSaveOptions UTF-8 output for Excel to JSON conversion | Specify character set for JSON output using Aspose.Cells Save method | Set encoding property on JsonSaveOptions before Workbook.Save in C#
// Tags: Aspose.Cells JsonSaveOptions character set | C# JSON export with UTF-8 using Aspose.Cells | Workbook.Save JSON encoding option | configure JSON output encoding in .NET | export Excel to JSON UTF-8 Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

// Creates a workbook, configures JsonSaveOptions to use UTF-8 encoding, ensures the output directory exists, and saves the workbook as a UTF-8 encoded JSON file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Configure JSON save options (SaveFormat is preset for JsonSaveOptions)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Determine output path and ensure its directory exists
            string outputPath = "output.json";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as JSON
            workbook.Save(outputPath, jsonOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
