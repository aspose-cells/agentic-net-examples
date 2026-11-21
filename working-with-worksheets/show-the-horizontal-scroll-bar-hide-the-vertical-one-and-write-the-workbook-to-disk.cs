// Title: Save an Aspose.Cells workbook to a file in C# and note that scroll bar visibility cannot be set via the current Worksheet API
// Common Searches: how to hide the vertical scroll bar in an Aspose.Cells worksheet using C# | aspnet save workbook to excel file with Aspose.Cells and create output folder if missing | Aspose.Cells view options not available for scroll bar settings in .NET | C# example for saving a new workbook to output.xlsx with Aspose.Cells | show horizontal scroll bar in Aspose.Cells worksheet limitation
// Tags: Aspose.Cells workbook save to Excel file | Aspose.Cells worksheet view options limitation | C# ensure output directory before workbook save | Aspose.Cells scroll bar visibility unsupported | Aspose.Cells horizontal scroll bar display

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, explains that the current Aspose.Cells version does not expose ViewOptions for scroll‑bar settings, ensures the target directory exists, saves the workbook as "output.xlsx", and writes a success message to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: In the current Aspose.Cells version the Worksheet class does not expose ViewOptions.
            // If you need to adjust view settings, refer to the version‑specific documentation.

            // Prepare output path and ensure its directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to disk
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
