using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (headers in row 1, data rows 2-6)
            // Column A will be used for grouping, column G (index 6) will have values for Min subtotal
            cells["A1"].PutValue("Group");
            cells["G1"].PutValue("Amount");

            // Sample data
            string[] groups = { "Alpha", "Alpha", "Beta", "Beta", "Gamma" };
            double[] amounts = { 120.5, 80.3, 150.0, 95.7, 200.2 };

            for (int i = 0; i < groups.Length; i++)
            {
                cells[i + 1, 0].PutValue(groups[i]);   // Column A
                cells[i + 1, 6].PutValue(amounts[i]); // Column G
            }

            // Define the cell area that contains the data (A1:G6)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = groups.Length,      // last data row index (0‑based)
                EndColumn = 6                // column G
            };

            // Apply subtotal:
            // - Group by column A (index 0)
            // - Use Min function
            // - Subtotal on column G (index 6)
            // - Replace existing subtotals, no page breaks, summary rows placed above data (flat view)
            cells.Subtotal(area, 0, ConsolidationFunction.Min, new int[] { 6 }, true, false, false);

            // Disable outline grouping for a flat view
            // Setting SummaryRowBelow to false ensures summary rows are not shown as grouped outlines
            worksheet.Outline.SummaryRowBelow = false;

            // Save the workbook
            workbook.Save("Subtotal_Min_ColumnG_FlatView.xlsx");
        }
    }
}