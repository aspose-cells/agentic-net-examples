// Title: Set worksheet default print resolution to 600 DPI with Aspose.Cells for .NET (C#)
// AI Prompts: Configure PageSetup.PrintQuality = 600 on a worksheet using Aspose.Cells in C#. | Programmatically increase Excel sheet print resolution to 600 DPI with the Aspose.Cells .NET API. | Apply high‑definition print quality to the first worksheet before saving the workbook.
// Common Searches: Aspose.Cells how to change default print DPI for a worksheet in C# | C# set Excel worksheet print quality to 600 DPI using Aspose.Cells | Increase print resolution of generated Excel file with Aspose.Cells .NET | Set PageSetup.PrintQuality property to 600 in Aspose.Cells example
// Tags: Aspose.Cells worksheet print quality DPI | C# PageSetup.PrintQuality property | high‑resolution Excel export Aspose.Cells | default print resolution .NET

using System;
using System.IO;
using Aspose.Cells;

// // Creates a workbook, accesses the first worksheet, sets PageSetup.PrintQuality to 600 DPI, ensures the output folder exists, and saves the file as HighResolutionOutput.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the default print quality (resolution) to 600 DPI
            sheet.PageSetup.PrintQuality = 600;

            // Define output file path
            string outputPath = "HighResolutionOutput.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
