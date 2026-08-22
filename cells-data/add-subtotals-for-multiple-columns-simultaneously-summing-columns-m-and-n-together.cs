// Title: Create grouped SUM subtotals for columns M and N in an Excel sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to group rows by column L and add SUM subtotals for both columns M and N in a single operation. | Show how to call Cells.Subtotal with a CellArea to subtotal multiple columns, insert page breaks, and place summary rows below each group in Aspose.Cells.
// Common Searches: aspocells c# subtotal multiple columns example | how to add sum subtotals for two columns in Aspose.Cells | group rows by a column and subtotal other columns using Aspose.Cells .NET | aspocells insert page breaks when creating subtotals | define CellArea for subtotal range Aspose.Cells C#
// Tags: Aspose.Cells Cells.Subtotal multiple columns | C# subtotal sum two columns Excel | group by column L Aspose.Cells | page break insertion Aspose.Cells subtotal | CellArea range for subtotal Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalMultipleColumns
{
    // The example creates a new workbook, fills columns L (group), M (Value1) and N (Value2) with sample data, defines a CellArea covering L1:N6, and calls Cells.Subtotal to group by column L, calculate SUM subtotals for columns M and N, replace existing subtotals, insert page breaks between groups, and place summary rows below each group. The workbook is saved as Subtotal_M_and_N.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Sample data setup
            // ------------------------------------------------------------
            // Header row (L = Group, M = Value1, N = Value2)
            cells["L1"].PutValue("Group");
            cells["M1"].PutValue("Value1");
            cells["N1"].PutValue("Value2");

            // Populate some rows with sample data
            // Group column (L) will be used for grouping
            string[] groups = { "A", "A", "B", "B", "C", "C" };
            double[,] values = {
                { 10, 20 },
                { 15, 25 },
                { 30, 40 },
                { 35, 45 },
                { 50, 60 },
                { 55, 65 }
            };

            for (int i = 0; i < groups.Length; i++)
            {
                cells[i + 1, 11].PutValue(groups[i]);          // Column L (index 11)
                cells[i + 1, 12].PutValue(values[i, 0]);      // Column M (index 12)
                cells[i + 1, 13].PutValue(values[i, 1]);      // Column N (index 13)
            }

            // ------------------------------------------------------------
            // Define the range that includes the header and data rows
            // Start at L1 (row 0, column 11) and end at N6 (row 5, column 13)
            // ------------------------------------------------------------
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 11,
                EndRow = groups.Length,   // includes header row + data rows
                EndColumn = 13
            };

            // ------------------------------------------------------------
            // Apply subtotals:
            // - Group by column L (zero‑based index 11)
            // - Use SUM function
            // - Add subtotals for both columns M (12) and N (13)
            // - Replace existing subtotals, add page breaks, place summary below data
            // ------------------------------------------------------------
            cells.Subtotal(
                area,
                11,                                 // groupBy column (L)
                ConsolidationFunction.Sum,          // subtotal function
                new int[] { 12, 13 },               // columns to subtotal (M and N)
                true,                               // replace existing subtotals
                true,                               // insert page breaks between groups
                true);                              // place summary rows below each group

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("Subtotal_M_and_N.xlsx");
        }
    }
}
