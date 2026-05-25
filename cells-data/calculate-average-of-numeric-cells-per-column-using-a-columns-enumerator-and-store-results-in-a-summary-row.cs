using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class AveragePerColumn
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ----- Populate sample data -----
        // Header row
        for (int c = 0; c < 4; c++)
            cells[0, c].PutValue($"Col{c + 1}");

        // Numeric data rows (rows 1..10)
        for (int r = 1; r <= 10; r++)
        {
            for (int c = 0; c < 4; c++)
            {
                // Example values: row index multiplied by column index (1‑based)
                cells[r, c].PutValue(r * (c + 1));
            }
        }

        // ----- Create a table covering the data range -----
        // Table range: from A1 (0,0) to D11 (10,3) – includes header + 10 data rows
        ListObjectCollection tables = worksheet.ListObjects;
        int tableIdx = tables.Add(0, 0, 10, 3, true);
        ListObject table = tables[tableIdx];

        // Show the totals row where averages will be placed
        table.ShowTotals = true;

        // ----- Set average calculation for each column using Columns enumerator -----
        foreach (Column col in cells.Columns)
        {
            int colIndex = col.Index;

            // Stop when we go beyond the last column of the table
            if (colIndex > 3) break;

            // The ListColumn at the same index corresponds to the table column
            table.ListColumns[colIndex].TotalsCalculation = TotalsCalculation.Average;
        }

        // ----- Save the workbook -----
        workbook.Save("AveragePerColumn.xlsx");
    }
}