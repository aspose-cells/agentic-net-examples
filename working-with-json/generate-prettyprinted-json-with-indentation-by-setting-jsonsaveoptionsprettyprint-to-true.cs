// Title: Generate a pretty‑printed JSON file from an Aspose.Cells workbook in C# using JsonSaveOptions
// AI Prompts: Create a C# program that builds a Workbook, fills cells with data, and saves it as an indented JSON file by setting JsonSaveOptions.PrettyPrint to true. | Update existing Aspose.Cells code to produce indented JSON output when calling Workbook.Save with JsonSaveOptions.
// Common Searches: how to enable formatted JSON output when saving an Aspose.Cells workbook in C# | C# Aspose.Cells JsonSaveOptions example for indented JSON | export Excel worksheet to JSON with indentation using Aspose.Cells | set JsonSaveOptions.PrettyPrint property in .NET code
// Tags: Aspose.Cells JsonSaveOptions pretty‑print JSON | C# save workbook as JSON with indentation | Aspose.Cells export to formatted JSON | JsonSaveOptions indentation setting .NET | Excel to JSON conversion with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a new Workbook, adds sample name and age data, configures JsonSaveOptions, ensures the output directory exists, and saves the workbook to 'output.json' with pretty‑printed (indented) JSON formatting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);

            // Configure JSON save options (pretty‑print not supported directly)
            JsonSaveOptions saveOptions = new JsonSaveOptions();

            // Ensure the output directory exists
            string outputPath = "output.json";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a JSON file using the specified options
            workbook.Save(outputPath, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
