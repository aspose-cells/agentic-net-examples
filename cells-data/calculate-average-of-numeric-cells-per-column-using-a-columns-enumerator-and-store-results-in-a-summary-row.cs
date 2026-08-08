// Title: Aspose.Cells C# – Compute Column Averages with a Totals Row via Columns Enumerator
// Description: Creates a workbook, populates a 5‑column × 10‑row numeric range, converts it to a ListObject (table), displays a totals row, and uses the worksheet Columns enumerator to set TotalsCalculation.Average for each column before saving the file.
// Keywords: Aspose.Cells column average | C# totals row average | Columns enumerator Aspose.Cells | ListObject totals calculation | Excel average per column C# | Aspose.Cells summary row | calculate column averages Aspose
// Common Searches: Aspose.Cells add totals row with average | set TotalsCalculation.Average for each column | C# iterate worksheet columns enumerator Aspose | how to compute column averages in Aspose.Cells | create summary row in Excel using Aspose.Cells
// Developer Intent: Add a totals row to a table and automatically calculate the average of numeric cells for every column.
// Use Cases: Financial reports that need an average sales row per month. | Dynamic dashboards where new data rows instantly update column averages. | Exporting data sets with a pre‑calculated average row for quick performance analysis.
// AI Prompts: Generate C# code with Aspose.Cells that creates a table, shows a totals row, and sets TotalsCalculation.Average for each column using the Columns enumerator. | Show how to iterate over worksheet columns in Aspose.Cells to apply average calculations to a ListObject's totals row. | Provide a complete example that fills sample numeric data, converts the range to a ListObject, enables a totals row, assigns average calculations, and saves the workbook.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, populates a 5‑column × 10‑row numeric range, converts it to a ListObject (table), displays a totals row, and uses the worksheet Columns enumerator to set TotalsCalculation.Average for each column before saving the file.
class AveragePerColumnWithTotalsRow
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ------------------------------------------------------------
        // Populate sample numeric data (5 columns, 10 data rows)
        // ------------------------------------------------------------
        int totalRows = 10;
        int totalCols = 5;

        // Header row
        for (int c = 0; c < totalCols; c++)
        {
            cells[0, c].PutValue($"Col{c + 1}");
        }

        // Data rows
        for (int r = 1; r <= totalRows; r++)
        {
            for (int c = 0; c < totalCols; c++)
            {
                // Example data: (row index) * (column index + 1)
                cells[r, c].PutValue(r * (c + 1));
            }
        }

        // ------------------------------------------------------------
        // Convert the range into a ListObject (table) so we can use a totals row
        // ------------------------------------------------------------
        int firstRow = 0;
        int firstCol = 0;
        int lastRow = totalRows;          // includes header row
        int lastCol = totalCols - 1;

        ListObjectCollection tables = sheet.ListObjects;
        int tableIndex = tables.Add(firstRow, firstCol, lastRow, lastCol, true);
        ListObject table = tables[tableIndex];

        // Show the totals row at the bottom of the table
        table.ShowTotals = true;

        // ------------------------------------------------------------
        // Use the Columns enumerator to set the TotalsCalculation for each column to Average
        // ------------------------------------------------------------
        IEnumerator colEnum = sheet.Cells.Columns.GetEnumerator();
        while (colEnum.MoveNext())
        {
            // Each item is a Column object
            Column col = (Column)colEnum.Current;

            // Column.Index gives the zero‑based column index
            int colIdx = col.Index;

            // Ensure the column is within the table range
            if (colIdx >= firstCol && colIdx <= lastCol)
            {
                // The ListColumns collection aligns with the table columns
                table.ListColumns[colIdx - firstCol].TotalsCalculation = TotalsCalculation.Average;
            }
        }

        // ------------------------------------------------------------
        // Save the workbook
        // ------------------------------------------------------------
        workbook.Save("AveragePerColumnWithTotalsRow.xlsx");
    }
}
