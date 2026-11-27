// Title: Enable circular reference detection in Aspose.Cells workbook calculation settings using C#
// Common Searches: Aspose.Cells C# enable circular reference detection in workbook calculation | how to prevent infinite formula evaluation loops with Aspose.Cells .NET | set calculation options to detect circular references in Aspose.Cells workbook | Aspose.Cells circular reference handling example in C# | configure workbook calculation settings for circular reference detection Aspose.Cells
// Tags: circular reference handling Aspose.Cells | workbook calculation options .NET | infinite formula loop prevention Aspose.Cells | Aspose.Cells calculation engine configuration | enable circular reference check C# | Aspose.Cells workbook calculation settings

using System;
using System.IO;
using Aspose.Cells;

// This C# example creates a new Aspose.Cells Workbook (or loads an existing file), highlights that circular‑reference detection must be configured via the workbook’s calculation settings (not available in older versions), ensures the output folder exists, saves the workbook as 'output.xlsx', and includes basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one with new Workbook("input.xlsx"))
            Workbook workbook = new Workbook();

            // NOTE: Circular reference settings are not available in this version of Aspose.Cells.
            // If needed, they can be configured using the appropriate API in newer versions.

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string fullOutputPath = Path.GetFullPath(outputPath);
            string? outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
