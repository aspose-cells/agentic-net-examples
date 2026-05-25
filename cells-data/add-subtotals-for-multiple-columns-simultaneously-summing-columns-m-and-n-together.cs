using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add header row (columns L, M, N)
        cells["L1"].PutValue("Group");
        cells["M1"].PutValue("Value1");
        cells["N1"].PutValue("Value2");

        // Sample data
        string[] groups = { "A", "A", "B", "B", "A" };
        int[,] values = {
            { 10, 20 },
            { 15, 25 },
            { 5, 30 },
            { 8, 12 },
            { 20, 10 }
        };

        // Populate data starting from row 2 (zero‑based index 1)
        for (int i = 0; i < groups.Length; i++)
        {
            cells[i + 1, 11].PutValue(groups[i]);          // Column L (index 11)
            cells[i + 1, 12].PutValue(values[i, 0]);      // Column M (index 12)
            cells[i + 1, 13].PutValue(values[i, 1]);      // Column N (index 13)
        }

        // Define the cell area that includes the header and data (L1:N6)
        CellArea area = CellArea.CreateCellArea(0, 11, groups.Length, 13);
        // Parameters:
        // - groupBy: 0 (first column of the area, i.e., column L)
        // - function: Sum
        // - totalList: {1, 2} (columns M and N within the area)
        // - replace, pageBreaks, summaryBelowData: true for demonstration
        worksheet.Cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1, 2 }, true, true, true);

        // Save the workbook
        workbook.Save("Subtotal_M_N.xlsx");
    }
}