using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

class TableToRangeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: 5 columns (A‑E) and 9 data rows (1‑9)
        for (int col = 0; col < 5; col++)
        {
            cells[0, col].PutValue($"Col{col + 1}"); // Header row
        }

        for (int row = 1; row <= 9; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                cells[row, col].PutValue(row * (col + 1));
            }
        }

        // Add a ListObject (table) that covers the populated area (A1:E10)
        int tableIndex = sheet.ListObjects.Add(0, 0, 9, 4, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Enable a totals row and set a sum calculation for the last column
        table.ShowTotals = true;
        table.ListColumns[4].TotalsCalculation = TotalsCalculation.Sum;

        // Configure conversion options – keep the same last row index (zero‑based)
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 9
        };

        // Convert the table to a normal range using the options
        table.ConvertToRange(options);

        // After conversion the table is removed; create a Range object for the same area
        // Total rows = data rows (9) + header (1) + totals row (1) = 11 rows, but ConvertToRange keeps the original range size (including totals)
        // Here we use 11 rows (0‑10) and 5 columns (0‑4)
        AsposeRange dataRange = cells.CreateRange(0, 0, 11, 5);

        // Example operation on the range: transpose rows and columns
        dataRange.Transpose();

        // Save the workbook to an XLSX file
        workbook.Save("TableConvertedToRange.xlsx");
    }
}