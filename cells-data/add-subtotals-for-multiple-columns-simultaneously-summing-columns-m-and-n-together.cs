using Aspose.Cells;
using System;

class SubtotalMultipleColumns
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add header row
        cells["A1"].PutValue("Group");   // Column A (index 0) – grouping field
        cells["M1"].PutValue("ValueM");  // Column M (index 12)
        cells["N1"].PutValue("ValueN");  // Column N (index 13)

        // Populate sample data (5 rows)
        for (int i = 0; i < 5; i++)
        {
            // Alternate groups for demonstration
            cells[i + 1, 0].PutValue(i % 2 == 0 ? "G1" : "G2"); // Group column
            cells[i + 1, 12].PutValue(10 + i);                 // Column M values
            cells[i + 1, 13].PutValue(20 + i);                 // Column N values
        }

        // Define the cell area covering the data (A1:N6)
        CellArea area = CellArea.CreateCellArea(0, 0, 5, 13);

        // Add subtotals:
        // - Group by column A (index 0)
        // - Use SUM function
        // - Apply to columns M and N (indices 12 and 13)
        // - Replace existing subtotals, no page breaks, summary below data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 12, 13 }, true, false, true);

        // Save the workbook
        workbook.Save("Subtotal_M_N.xlsx");
    }
}