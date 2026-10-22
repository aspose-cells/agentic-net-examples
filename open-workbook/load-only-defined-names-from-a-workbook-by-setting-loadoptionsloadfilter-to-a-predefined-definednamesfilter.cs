// Title: Read all defined names (named ranges) from an Excel workbook using Aspose.Cells in C#
// Common Searches: how to list named ranges in an Excel file with Aspose.Cells C# | Aspose.Cells C# enumerate defined names without modifying the workbook | retrieve all Excel defined names using Aspose.Cells .NET API | C# code to get named ranges from an XLSX file with Aspose.Cells | extract workbook defined names Aspose.Cells example in C#
// Tags: Aspose.Cells read defined names C# | Aspose.Cells enumerate named ranges .NET | LoadOptions workbook loading C# | Aspose.Cells workbook named ranges extraction | C# iterate Excel defined names Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input XLSX file, loads the workbook with default LoadOptions, iterates through the workbook's defined names (named ranges), prints each name and its reference, ensures the output directory exists, and saves the workbook unchanged.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (no filter applied to keep compatibility with all versions)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through defined names and display their details
            foreach (var definedName in workbook.Worksheets.Names)
            {
                // Use Text property for the name (compatible with older Aspose.Cells versions)
                string name = definedName.Text ?? string.Empty;
                string refersTo = definedName.RefersTo ?? string.Empty;
                Console.WriteLine($"Name: {name}, RefersTo: {refersTo}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional, as no modifications are made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
