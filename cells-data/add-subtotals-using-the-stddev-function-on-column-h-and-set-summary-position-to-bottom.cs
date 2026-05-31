using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data: column A as the grouping field, column H as the values for StdDev
        cells["A1"].PutValue("Group");
        cells["H1"].PutValue("Value");

        string[] groups = { "X", "X", "Y", "Y", "Z" };
        double[] values = { 10, 20, 30, 40, 50 };

        for (int i = 0; i < groups.Length; i++)
        {
            cells[i + 1, 0].PutValue(groups[i]);   // Column A (index 0)
            cells[i + 1, 7].PutValue(values[i]);   // Column H (index 7)
        }

        // Define the cell area covering the data (from A1 to H{rows})
        CellArea area = CellArea.CreateCellArea(0, 0, groups.Length, 7);

        // Add subtotals:
        // - Group by column A (index 0)
        // - Use StdDev function on column H (index 7)
        // - Do not replace existing subtotals, no page breaks, summary placed below data (bottom)
        cells.Subtotal(area, 0, ConsolidationFunction.StdDev, new int[] { 7 }, false, false, true);

        // Save the workbook
        workbook.Save("SubtotalStdDevBottom.xlsx");
    }
}