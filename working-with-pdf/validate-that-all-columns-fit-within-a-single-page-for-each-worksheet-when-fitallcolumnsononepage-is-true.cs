// Title: How to validate that all columns fit on a single printed page when FitToPagesWide = 1 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells check if worksheet columns fit on one printed page in C# | C# calculate printable width of Excel sheet with margins using Aspose.Cells | Validate column width against page setup printable area Aspose.Cells .NET | Determine if FitToPagesWide setting causes column overflow in Aspose.Cells | Get column width in inches with Aspose.Cells Cells.GetColumnWidth
// Tags: FitToPagesWide overflow detection | printable page width calculation Aspose.Cells | Cells.GetColumnWidth unit conversion | worksheet page setup printable area validation | Aspose.Cells column fit single page

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates each worksheet, and when PageSetup.FitToPagesWide equals 1 (FitAllColumnsOnOnePage), it computes the printable page width by subtracting left and right margins from the paper width. It then sums the widths of all used columns in inches using Cells.GetColumnWidth, compares the total to the printable width, and logs whether the columns fit on a single printed page before optionally saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;

                // Check if the worksheet is set to fit all columns on one page
                // In Aspose.Cells this is indicated by FitToPagesWide == 1
                if (pageSetup.FitToPagesWide == 1)
                {
                    // Calculate printable page width (in inches)
                    double printableWidth = pageSetup.PaperWidth - (pageSetup.LeftMargin + pageSetup.RightMargin);

                    // Calculate total width of all used columns (in inches)
                    double totalColumnWidth = 0.0;
                    int maxColumn = sheet.Cells.MaxColumn; // zero‑based index of the last column with data
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Use the recommended API to get column width in inches
                        totalColumnWidth += sheet.Cells.GetColumnWidth(col, true, CellsUnitType.Inch);
                    }

                    // Validate whether columns fit within the printable area
                    if (totalColumnWidth > printableWidth)
                    {
                        Console.WriteLine($"Worksheet '{sheet.Name}' columns exceed printable width: {totalColumnWidth:F2}\" > {printableWidth:F2}\".");
                    }
                    else
                    {
                        Console.WriteLine($"Worksheet '{sheet.Name}' columns fit within a single page.");
                    }
                }
            }

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
