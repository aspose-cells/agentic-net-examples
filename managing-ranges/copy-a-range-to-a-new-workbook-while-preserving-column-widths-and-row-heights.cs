// Title: Copy a cell range to a new workbook while preserving column widths and row heights with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy a range to another workbook keeping original column widths | how to retain row heights when copying Excel cells to a new file using Aspose.Cells | sample code for transferring a selected range and its formatting between workbooks in .NET | preserve column width and row height during range copy with Aspose.Cells API | copy A1:D10 range to new workbook and maintain layout using Aspose.Cells for C#
// Tags: copy range preserving column width Aspose.Cells | retain row height during range copy .NET | transfer column widths between workbooks C# | copy worksheet range to new workbook Aspose.Cells | preserve formatting while copying Excel range .NET | set column width after range copy Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a source Excel workbook, defines a specific cell range, copies its values, formulas, formats, and merged cells to a new workbook, then iterates over the range's columns and rows to replicate the original column widths and row heights before saving the destination file.
class Program
{
    static void Main()
    {
        try
        {
            const string srcPath = "source.xlsx";
            const string destPath = "destination.xlsx";

            // Verify source file exists
            if (!File.Exists(srcPath))
            {
                Console.WriteLine($"Source file '{srcPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(srcPath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Define the range to copy (example: A1:D10)
            int startRow = 0;          // Row 1 (zero‑based)
            int startColumn = 0;       // Column A
            int rowCount = 10;         // Number of rows to copy
            int columnCount = 4;       // Number of columns to copy

            // Create the destination workbook
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Copy cell values, formulas, formats, and merged cells
            Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount);
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount);
            srcRange.Copy(destRange);

            // Preserve column widths
            for (int col = startColumn; col < startColumn + columnCount; col++)
            {
                double width = srcSheet.Cells.GetColumnWidth(col);
                destSheet.Cells.SetColumnWidth(col, width);
            }

            // Preserve row heights
            for (int row = startRow; row < startRow + rowCount; row++)
            {
                double height = srcSheet.Cells.GetRowHeight(row);
                destSheet.Cells.SetRowHeight(row, height);
            }

            // Save the new workbook
            destWorkbook.Save(destPath);
            Console.WriteLine($"Destination workbook saved as '{destPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
