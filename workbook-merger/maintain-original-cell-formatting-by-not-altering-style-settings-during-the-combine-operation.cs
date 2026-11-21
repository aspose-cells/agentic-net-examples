// Title: Combine two Excel worksheets into one workbook while preserving cell styles using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge two workbooks keep original cell formatting | copy rows between worksheets without losing styles Aspose.Cells | how to combine Excel sheets preserving formatting with Aspose.Cells .NET | merge multiple Excel files while retaining cell styles using Aspose.Cells | preserve cell style when appending worksheets in Aspose.Cells C#
// Tags: workbook merge preserving cell styles Aspose.Cells | copy rows retain formatting Aspose.Cells | combine worksheets without losing styles .NET | Excel file concatenation with formatting retention | cells.copyrows keep cell formatting C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads two source Excel workbooks, creates a new destination workbook, clears its default sheet, adds a "Combined" worksheet, and copies each row from the source worksheets to the destination using Cells.CopyRows, which retains values, formulas, and style settings. A blank row separates the sections, and the merged workbook is saved as Combined.xlsx with all original formatting intact.
class Program
{
    static void Main()
    {
        try
        {
            string srcPath1 = "Source1.xlsx";
            string srcPath2 = "Source2.xlsx";

            // Verify that source files exist before loading
            if (!File.Exists(srcPath1))
                throw new FileNotFoundException($"Source file not found: {srcPath1}");
            if (!File.Exists(srcPath2))
                throw new FileNotFoundException($"Source file not found: {srcPath2}");

            // Load source workbooks (preserves all original formatting)
            Workbook srcWb1 = new Workbook(srcPath1);
            Workbook srcWb2 = new Workbook(srcPath2);

            // Create a new workbook for the combined result
            Workbook destWb = new Workbook();
            // Remove the default sheet that Aspose.Cells creates
            destWb.Worksheets.Clear();

            // Add a worksheet that will hold the combined data
            Worksheet destSheet = destWb.Worksheets.Add("Combined");

            int destRow = 0; // tracks the next free row in the destination sheet

            // Helper to copy all rows from a source worksheet to the destination,
            // using CopyRows which retains cell values, formulas, and style settings.
            void CopyWorksheetRows(Worksheet srcSheet)
            {
                int maxRow = srcSheet.Cells.MaxDataRow;
                for (int r = 0; r <= maxRow; r++)
                {
                    destSheet.Cells.CopyRows(srcSheet.Cells, r, destRow, 1);
                    destRow++; // advance to the next free row in the destination
                }
            }

            // Copy the first worksheet's rows (preserving formatting)
            CopyWorksheetRows(srcWb1.Worksheets[0]);

            // Optional: add a blank row between combined sections
            destRow++;

            // Copy the second worksheet's rows (preserving formatting)
            CopyWorksheetRows(srcWb2.Worksheets[0]);

            // Save the combined workbook (original formatting remains intact)
            string outputPath = "Combined.xlsx";
            destWb.Save(outputPath);
            Console.WriteLine($"Combined workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
