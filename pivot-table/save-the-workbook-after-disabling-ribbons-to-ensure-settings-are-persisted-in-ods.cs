// Title: Saving an Aspose.Cells workbook as ODS in C# while handling the unavailable ribbon‑disabling feature
// Common Searches: how to export a workbook to ODS format using Aspose.Cells for .NET | does Aspose.Cells support disabling ribbons when saving to ODS | persist ribbon visibility settings in an ODS file with Aspose.Cells C# | Aspose.Cells SaveFormat.Ods example code C# | workaround for missing DisableRibbons property in Aspose.Cells
// Tags: Aspose.Cells save workbook as ODS C# | disable ribbons not supported Aspose.Cells | persist UI settings in ODS export | SaveFormat.Ods usage Aspose.Cells | C# OpenDocument Spreadsheet export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Aspose.Cells Workbook, notes that the DisableRibbons property is unavailable in the current version, and saves the workbook to an OpenDocument Spreadsheet (Result.ods) using SaveFormat.Ods, with exception handling and path output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // NOTE: The DisableRibbons property is not available in the current Aspose.Cells version.
            // If ribbon disabling is required, refer to the latest documentation for the appropriate API.

            // Define output file path
            string outputPath = "Result.ods";

            // Save the workbook as ODS (using the non‑obsolete enum value)
            workbook.Save(outputPath, SaveFormat.Ods);

            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
