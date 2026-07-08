using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalVarExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including column I which is index 8)
            // Header row
            cells["A1"].PutValue("Group");
            cells["I1"].PutValue("Value");

            // Sample data rows
            // Group column (A) will be used for grouping
            // Column I will contain numeric values for which we calculate variance
            string[] groups = { "Alpha", "Alpha", "Beta", "Beta", "Beta", "Gamma", "Gamma" };
            double[] values = { 10, 12, 20, 22, 24, 30, 32 };

            for (int i = 0; i < groups.Length; i++)
            {
                cells[i + 1, 0].PutValue(groups[i]);   // Column A (index 0)
                cells[i + 1, 8].PutValue(values[i]);  // Column I (index 8)
            }

            // Define the range that includes the header and all data rows
            // From A1 to I{lastRow}
            int lastRow = groups.Length; // because rows are 1-based after header
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = lastRow,
                EndColumn = 8
            };

            // Add subtotals:
            // - Group by column A (index 0)
            // - Use Var function (variance)
            // - Apply subtotal to column I (index 8)
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(area, 0, ConsolidationFunction.Var, new int[] { 8 }, true, true, true);

            // Enable outline display (summary rows positioned below the detail rows)
            sheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("SubtotalVarOutline.xlsx");
        }
    }
}