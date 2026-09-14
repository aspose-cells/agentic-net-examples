// Title: Check if all columns fit on one printed page for worksheets with FitAllColumnsOnOnePage enabled using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through each worksheet in a workbook and reports whether the combined width of used columns fits within the printable area when FitToPagesWide is set to 1. | Create a reusable function in Aspose.Cells that returns a boolean indicating if a worksheet's columns can be printed on a single page, handling pixel‑to‑point conversion and margin calculations. | Write a C# routine that lists the column indexes that cause the total width to exceed the printable page width for worksheets with FitAllColumnsOnOnePage enabled.
// Common Searches: how to verify column width fits on one page with Aspose.Cells C# | Aspose.Cells check FitAllColumnsOnOnePage column overflow | C# calculate printable page width for Excel worksheet using Aspose.Cells | determine if Excel columns exceed printable area when FitToPagesWide = 1 | validate worksheet column layout before PDF export Aspose.Cells
// Tags: Aspose.Cells column width validation | FitAllColumnsOnOnePage printable width check | pixel to point conversion Aspose.Cells | worksheet margin calculation Aspose.Cells | Excel to PDF column overflow detection

using System;
using System.IO;
using Aspose.Cells;

// C# program that loads an XLSX workbook, examines each worksheet with FitToPagesWide set to 1, converts column pixel widths to points, computes the printable page width from margins, and reports whether the total column width fits on a single printed page.
class ColumnFitValidator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the worksheet is configured to fit all columns on one page
                // In Aspose.Cells this is represented by FitToPagesWide == 1
                bool fitAllColumnsOnOnePage = sheet.PageSetup.FitToPagesWide == 1;

                if (fitAllColumnsOnOnePage)
                {
                    // Calculate printable width of the page (points)
                    // PaperWidth, LeftMargin, RightMargin are in points (1 point = 1/72 inch)
                    double printableWidth = sheet.PageSetup.PaperWidth
                                            - sheet.PageSetup.LeftMargin
                                            - sheet.PageSetup.RightMargin;

                    // Sum the widths of all used columns
                    double totalColumnWidthPoints = 0.0;
                    int maxColumn = sheet.Cells.MaxDataColumn; // last column that contains data

                    // If the sheet is empty, skip validation
                    if (maxColumn < 0)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\": empty sheet, nothing to validate.");
                        continue;
                    }

                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Get column width in pixels via Cells collection
                        int widthPixels = sheet.Cells.GetColumnWidthPixel(col);
                        // Convert pixels to points (default screen DPI = 96)
                        double widthPoints = widthPixels * 72.0 / 96.0; // 1 point = 1/72 inch
                        totalColumnWidthPoints += widthPoints;
                    }

                    // Validate whether total column width fits within printable width
                    if (totalColumnWidthPoints <= printableWidth)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\": all columns fit on one page (total width {totalColumnWidthPoints:F2} pt, printable width {printableWidth:F2} pt).");
                    }
                    else
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\": columns exceed one page (total width {totalColumnWidthPoints:F2} pt, printable width {printableWidth:F2} pt).");
                    }
                }
                else
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\": FitAllColumnsOnOnePage is disabled, validation skipped.");
                }
            }

            // No changes are made to the workbook, so saving is optional.
            // Uncomment the following line if you need to save the workbook.
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
