// Title: Copy a specific cell range to a new workbook while preserving column widths and row heights using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy range to another workbook keep column width and row height | How to duplicate a cell block with formatting and formulas preserving dimensions in .NET | Copy Excel range to new file preserving column pixel width using Aspose.Cells | Transfer A1:D10 range to a new workbook while maintaining row heights in C# | Preserve column widths when copying cells between workbooks with Aspose.Cells API
// Tags: copy range to new workbook Aspose.Cells | preserve column width pixel Aspose.Cells | preserve row height pixel Aspose.Cells | duplicate cell styles and formulas Aspose.Cells | copy Excel range programmatically .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads source.xlsx, copies the A1:D10 range (including values, formulas, and styles) to a newly created workbook, replicates the original column widths and row heights using GetColumnWidthPixel/SetColumnWidthPixel and GetRowHeightPixel/SetRowHeightPixel, and saves the result as destination.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(sourcePath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Define the range to copy (example: A1:D10)
            int startRow = 0;      // A1 -> row 0
            int startColumn = 0;   // A1 -> column 0
            int endRow = 9;        // row 10 (zero‑based)
            int endColumn = 3;     // column D (zero‑based)

            // Create a new workbook for the destination
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Copy cells (values, formulas, styles) manually
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell srcCell = srcSheet.Cells[row, col];
                    Cell destCell = destSheet.Cells[row, col];

                    // Copy value
                    destCell.PutValue(srcCell.Value);

                    // Copy formula if present
                    if (srcCell.IsFormula)
                    {
                        destCell.Formula = srcCell.Formula;
                    }

                    // Copy style (formatting)
                    Style srcStyle = srcCell.GetStyle();
                    destCell.SetStyle(srcStyle);
                }
            }

            // Preserve column widths
            for (int col = startColumn; col <= endColumn; col++)
            {
                int width = srcSheet.Cells.GetColumnWidthPixel(col);
                destSheet.Cells.SetColumnWidthPixel(col, width);
            }

            // Preserve row heights
            for (int row = startRow; row <= endRow; row++)
            {
                int height = srcSheet.Cells.GetRowHeightPixel(row);
                destSheet.Cells.SetRowHeightPixel(row, height);
            }

            // Save the new workbook
            destWorkbook.Save(destinationPath);
            Console.WriteLine($"Range copied successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
