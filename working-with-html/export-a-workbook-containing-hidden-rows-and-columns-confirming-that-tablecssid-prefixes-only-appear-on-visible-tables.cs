// Title: Export an Excel workbook with hidden rows and columns and list only visible tables using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells hide a specific row and column before saving workbook in C# | How to retrieve only visible ListObject tables after hiding rows with Aspose.Cells .NET | C# code to export Excel file with hidden rows/columns and check visible tables using Aspose.Cells | Determine if the first data row of a ListObject is hidden with Aspose.Cells API | Aspose.Cells HTML export visible tables only TableCssId prefix handling
// Tags: Aspose.Cells hide rows and columns | Aspose.Cells filter visible ListObjects | Aspose.Cells export workbook with hidden rows | Aspose.Cells TableStyle usage for HTML export | Aspose.Cells check first data row visibility

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, fills a 5x3 range, adds a styled ListObject table, hides the third row and second column, iterates through all tables to output only those whose first data row is not hidden, and saves the workbook as HiddenRowsColumnsWorkbook.xlsx.
class ExportWorkbookWithHiddenRowsColumns
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data (5 rows x 3 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Add a table (ListObject) covering the data range A1:C5
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 5;
            int totalColumns = 3;

            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];

            // Assign a table style name (used for HTML export styling)
            table.TableStyleName = "TableStyleMedium2";

            // Hide a row (row 3) and a column (column B)
            sheet.Cells.Rows[2].IsHidden = true;      // hides the third row (index 2)
            sheet.Cells.Columns[1].IsHidden = true;   // hides the second column (index 1)

            // Verify that tables whose first data row is hidden are treated as not visible
            Console.WriteLine("Visible tables and their style names:");
            foreach (ListObject lo in sheet.ListObjects)
            {
                // Determine if the table's first data row is hidden
                int firstDataRow = lo.DataRange.FirstRow;
                bool isFirstRowHidden = sheet.Cells.Rows[firstDataRow].IsHidden;
                if (!isFirstRowHidden)
                {
                    // Table is considered visible; output its display name and style
                    Console.WriteLine($"Table DisplayName: {lo.DisplayName}, Style: {lo.TableStyleName}");
                }
            }

            // Save the workbook to a file (includes hidden rows/columns)
            string outputPath = "HiddenRowsColumnsWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
