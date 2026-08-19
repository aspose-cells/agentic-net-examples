// Title: Aspose.Cells .NET: Generate a report of cells where Style.QuotePrefix is true (row & column indices)
// Description: Creates a workbook, marks selected cells with Style.QuotePrefix, scans the used range, lists every cell whose QuotePrefix flag is true, and outputs the zero‑based row, column and A1 address to the console and a new worksheet before saving the file.
// Keywords: Aspose.Cells | .NET | C# | QuotePrefix | Style.QuotePrefix | list prefixed cells | find cells with quote prefix | row index | column index | used range iteration | generate report | export to worksheet
// Common Searches: Aspose.Cells find cells with QuotePrefix true | list cells that have a leading apostrophe using Aspose.Cells | C# report of QuotePrefix cells with row and column numbers | how to iterate used range for Style.QuotePrefix in Aspose.Cells | save prefixed‑cell report to new worksheet Aspose.Cells
// Developer Intent: Identify every cell whose Style.QuotePrefix property is set to true and produce a concise report that includes its zero‑based row, column, and cell address.
// Use Cases: Audit spreadsheets for cells that start with an apostrophe to prevent data‑type errors. | Export a summary of prefixed cells for data‑cleansing or validation workflows. | Create documentation that lists all QuotePrefix cells for regulatory or review purposes.
// AI Prompts: Write C# code with Aspose.Cells that scans a worksheet and returns the row, column and A1 address of all cells where Style.QuotePrefix is true. | Provide a more efficient version of the QuotePrefix report using Cells.Find or style enumeration. | Show how to modify the example to output the report as CSV and include only the A1 cell addresses.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Creates a workbook, marks selected cells with Style.QuotePrefix, scans the used range, lists every cell whose QuotePrefix flag is true, and outputs the zero‑based row, column and A1 address to the console and a new worksheet before saving the file.
class QuotePrefixReport
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // -------------------------------------------------
        // Sample data: set some cells with QuotePrefix = true
        // -------------------------------------------------
        Cell cellA1 = cells["A1"];
        cellA1.PutValue("'Sample with prefix");
        Style styleA1 = cellA1.GetStyle();
        styleA1.QuotePrefix = true;
        cellA1.SetStyle(styleA1);

        Cell cellB2 = cells["B2"];
        cellB2.PutValue("Normal text"); // No QuotePrefix

        Cell cellC3 = cells["C3"];
        cellC3.PutValue("'Another prefixed");
        Style styleC3 = cellC3.GetStyle();
        styleC3.QuotePrefix = true;
        cellC3.SetStyle(styleC3);

        // -------------------------------------------------
        // Build a report of cells where QuotePrefix is true
        // -------------------------------------------------
        List<string> reportLines = new List<string>();
        reportLines.Add("Cells with QuotePrefix = true (Row, Column, CellName):");

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell == null) continue;

                Style style = cell.GetStyle();
                if (style != null && style.QuotePrefix)
                {
                    // Row and Column indices are zero‑based
                    reportLines.Add($"Row {row}, Column {col}, Cell {cell.Name}");
                }
            }
        }

        // -------------------------------------------------
        // Output the report to the console
        // -------------------------------------------------
        foreach (string line in reportLines)
        {
            Console.WriteLine(line);
        }

        // -------------------------------------------------
        // Write the report into a new worksheet for reference
        // -------------------------------------------------
        int reportSheetIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
        for (int i = 0; i < reportLines.Count; i++)
        {
            reportSheet.Cells[i, 0].PutValue(reportLines[i]);
        }

        // -------------------------------------------------
        // Save the workbook containing the original data and the report
        // -------------------------------------------------
        workbook.Save("QuotePrefixReport.xlsx");
    }
}
