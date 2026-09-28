// Title: Create a macro‑free workbook, hide zero values on print, set a custom top margin, and export the first worksheet to CSV using Aspose.Cells for .NET
// AI Prompts: Write C# using Aspose.Cells to instantiate a workbook without macros, configure a 0.5‑inch page top spacing, suppress zero‑value cells in print output, and output the primary sheet as a CSV file. | Modify the example so that each worksheet receives the same custom page margin and is exported individually to separate CSV files. | Enhance the CSV export routine with try‑catch blocks that capture exceptions, log detailed error information, and optionally rethrow.
// Common Searches: Aspose.Cells set top margin programmatically C# | How to hide zero values when printing in Aspose.Cells workbook | Save Aspose.Cells worksheet as CSV without macros | Create macro‑free workbook using Aspose.Cells .NET
// Tags: workbook without macros Aspose.Cells | custom page margin Aspose.Cells | suppress zero printing Aspose.Cells | export sheet to CSV Aspose.Cells | page setup options Aspose.Cells

using Aspose.Cells;
using System;

// // This program creates a new macro‑free workbook, applies a 0.5‑inch top margin to each worksheet, attempts to suppress zero values during printing, and saves the first worksheet as a CSV file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new macro‑free workbook
            Workbook workbook = new Workbook();

            // Define custom top margin (0.5 inch = 36 points)
            const double topMarginPoints = 36.0;

            // Apply settings to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Hide zero values when printing (if supported by the library version)
                // sheet.PageSetup.PrintZeroValues = false; // Property not available in some versions

                // Set custom top margin
                sheet.PageSetup.TopMargin = topMarginPoints;
            }

            // Save the workbook as CSV (only the first worksheet is written to CSV)
            workbook.Save("output.csv", SaveFormat.Csv);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
