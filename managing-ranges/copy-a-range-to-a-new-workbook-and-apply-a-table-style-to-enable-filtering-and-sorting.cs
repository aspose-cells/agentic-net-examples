// Title: Copy a specific cell range to a new workbook and add a styled ListObject table with filtering in Aspose.Cells for .NET
// AI Prompts: Generate C# code that extracts cells A1:D10 from a source workbook, writes them into a fresh workbook while keeping all styles, and then defines a ListObject covering the transferred cells. | Show how to set TableStyleMedium9 on the ListObject so that Excel’s filter and sort dropdowns are displayed in the saved file. | Add code to assign a custom name to the ListObject and enable the auto‑filter feature after the range has been transferred using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells copy range to another workbook and create Excel table with filter | C# how to add ListObject to a transferred range using Aspose.Cells | Apply built‑in table style to a programmatically created table in Aspose.Cells .NET | Preserve cell formatting when moving cells between workbooks with Aspose.Cells | Set table name and enable auto‑filter for ListObject in Aspose.Cells C#
// Tags: copy range to new workbook Aspose.Cells | create ListObject table programmatically .NET | apply built-in table style Aspose.Cells | preserve cell formatting Aspose.Cells copy | enable filter sorting Aspose.Cells table

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads source.xlsx, copies the A1:D10 range (values and styles) into a fresh workbook, creates a ListObject over the copied area with a header row, applies the built‑in TableStyleMedium9 to provide filter and sort UI, and saves the result as destination.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "destination.xlsx";

            // Ensure the source file exists before loading
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(sourcePath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Define the source range to copy (A1:D10)
            CellArea srcArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 9,        // Row 10
                EndColumn = 3      // Column D
            };

            // Create a new workbook for the destination
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Copy cell values and styles within the defined area
            for (int r = srcArea.StartRow; r <= srcArea.EndRow; r++)
            {
                for (int c = srcArea.StartColumn; c <= srcArea.EndColumn; c++)
                {
                    // Copy value
                    destSheet.Cells[r, c].PutValue(srcSheet.Cells[r, c].Value);

                    // Copy style
                    Style style = srcSheet.Cells[r, c].GetStyle();
                    destSheet.Cells[r, c].SetStyle(style);
                }
            }

            // Calculate dimensions of the copied range
            int firstRow = srcArea.StartRow;
            int firstCol = srcArea.StartColumn;
            int totalRows = srcArea.EndRow - srcArea.StartRow + 1;
            int totalCols = srcArea.EndColumn - srcArea.StartColumn + 1;

            // Add a table (ListObject) over the copied range; assume the first row contains headers
            int tableIndex = destSheet.ListObjects.Add(
                firstRow,
                firstCol,
                firstRow + totalRows - 1,
                firstCol + totalCols - 1,
                true); // true = has header row

            ListObject table = destSheet.ListObjects[tableIndex];

            // Apply a built‑in table style to enable filtering and sorting UI
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the new workbook
            destWorkbook.Save(destPath);
            Console.WriteLine($"Workbook saved successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
