// Title: C# – Add Subtotals for Columns M and N Grouped by Column L with Aspose.Cells
// Description: Creates a workbook, fills columns L (group), M and N (values), defines the range L1:N6, and uses Cells.Subtotal with the SUM function to calculate subtotals for both M and N per group. The file is saved as SubtotalMultipleColumns.xlsx.
// Keywords: Aspose.Cells subtotal multiple columns | C# Cells.Subtotal example | group by column L Aspose.Cells | sum columns M N Aspose.Cells | subtotal function .NET | Excel subtotal automation C#
// Common Searches: Aspose.Cells add subtotals for several columns | C# subtotal multiple columns Aspose.Cells | group by column and sum other columns Aspose.Cells | how to use Cells.Subtotal in .NET | subtotal rows by category Aspose.Cells
// Developer Intent: Generate a worksheet, populate group and numeric data, and apply a single subtotal operation that sums two adjacent columns for each group.
// Use Cases: Sales report: subtotal quantity (M) and revenue (N) for each region (L). | Inventory summary: subtotal on‑hand and back‑order counts by product category. | Financial ledger: subtotal debit and credit amounts per account type.
// AI Prompts: Write C# code with Aspose.Cells to subtotal columns D and E grouped by column C. | Explain the totalList parameter in Cells.Subtotal and how to add more columns to the subtotal. | Extend the example to also compute the average of column M for each group.

using Aspose.Cells;
using System;

// Creates a workbook, fills columns L (group), M and N (values), defines the range L1:N6, and uses Cells.Subtotal with the SUM function to calculate subtotals for both M and N per group. The file is saved as SubtotalMultipleColumns.xlsx.
class SubtotalMultipleColumns
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Header row (columns L, M, N)
        cells["L1"].PutValue("Group");   // Column L (index 11)
        cells["M1"].PutValue("ValueM");  // Column M (index 12)
        cells["N1"].PutValue("ValueN");  // Column N (index 13)

        // Sample data
        string[] groups = { "A", "A", "B", "B", "A" };
        int[,] values = {
            { 10, 20 },
            { 15, 25 },
            { 5, 30 },
            { 8, 12 },
            { 20, 10 }
        };

        // Populate rows starting at row 2 (zero‑based index 1)
        for (int i = 0; i < groups.Length; i++)
        {
            cells[i + 1, 11].PutValue(groups[i]);          // Column L
            cells[i + 1, 12].PutValue(values[i, 0]);      // Column M
            cells[i + 1, 13].PutValue(values[i, 1]);      // Column N
        }

        // Define the range that includes the header and all data rows (L1:N6)
        // StartRow = 0, StartColumn = 11 (L), EndRow = groups.Length, EndColumn = 13 (N)
        CellArea area = CellArea.CreateCellArea(0, 11, groups.Length, 13);

        // Apply subtotals:
        // - Group by the first column of the area (Group column L)
        // - Use SUM function
        // - Add subtotals for the second and third columns of the area (M and N)
        // - Replace existing subtotals, insert page breaks, and place summary below data
        cells.Subtotal(
            area,
            0,                                 // groupBy offset within the area
            ConsolidationFunction.Sum,
            new int[] { 1, 2 },                // totalList offsets for M and N
            true,                              // replace existing subtotals
            true,                              // add page breaks between groups
            true                               // place summary below data
        );

        // Save the workbook
        workbook.Save("SubtotalMultipleColumns.xlsx");
    }
}
