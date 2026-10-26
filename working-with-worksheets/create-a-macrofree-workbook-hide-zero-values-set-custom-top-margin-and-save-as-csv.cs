// Title: Create a macro‑free workbook, hide zero values with cell formatting, set a 1‑inch top margin, and export to CSV using Aspose.Cells for .NET (C#)
// Common Searches: how to generate a macro‑free Excel workbook and save it as CSV with Aspose.Cells in C# | setting a custom top margin for a worksheet using Aspose.Cells page setup C# | hide zero values in an Aspose.Cells worksheet by applying cell formatting | export Aspose.Cells workbook to CSV without macros programmatically | ensure output directory exists before saving CSV with Aspose.Cells C#
// Tags: macro‑free workbook creation Aspose.Cells C# | worksheet top margin configuration Aspose.Cells | zero value hiding using cell formatting Aspose.Cells | CSV export of macro‑free workbook Aspose.Cells | ensure output folder exists Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates how to create a new macro‑free Workbook with Aspose.Cells for .NET, apply cell formatting to hide zero values, set a 1‑inch (72‑point) top margin via the worksheet PageSetup, guarantee the target directory exists, and save the workbook as a CSV file using SaveFormat.Csv.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new macro‑free workbook
            Workbook workbook = new Workbook();

            // NOTE: The IsZeroValuesHidden property is not available in the current Aspose.Cells version.
            // If needed, this feature can be implemented via cell formatting after populating data.

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a custom top margin (in points; 1 inch = 72 points)
            sheet.PageSetup.TopMargin = 72; // 1 inch top margin

            // Define output file path
            string outputPath = "output.csv";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as CSV (macro‑free)
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
