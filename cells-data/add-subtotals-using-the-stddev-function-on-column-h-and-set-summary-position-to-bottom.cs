using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalStdDevDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data with a grouping column (G) and a numeric column (H)
            // Header row
            cells["G1"].PutValue("Group");
            cells["H1"].PutValue("Value");

            // Sample data rows
            string[] groups = { "A", "A", "B", "B", "A", "B" };
            double[] values = { 10, 12, 15, 18, 11, 16 };

            for (int i = 0; i < groups.Length; i++)
            {
                cells[i + 1, 6].PutValue(groups[i]);   // Column G (index 6)
                cells[i + 1, 7].PutValue(values[i]);   // Column H (index 7)
            }

            // Define the cell area that includes the header and data (G1:H7)
            CellArea area = CellArea.CreateCellArea("G1", "H7");

            // Add subtotals:
            // - Group by column G (zero‑based index 6)
            // - Use StdDev function (ConsolidationFunction.StdDev)
            // - Apply subtotal to column H (zero‑based index 7)
            // - Do not replace existing subtotals, no page breaks, summary placed below data (bottom)
            worksheet.Cells.Subtotal(
                area,
                6,                                 // groupBy column index (G)
                ConsolidationFunction.StdDev,     // StdDev function
                new int[] { 7 },                  // totalList: column H
                false,                            // replace existing subtotals
                false,                            // add page breaks between groups
                true);                            // summaryBelowData = true (bottom)

            // Save the workbook
            workbook.Save("SubtotalStdDevBottom.xlsx");
        }
    }
}