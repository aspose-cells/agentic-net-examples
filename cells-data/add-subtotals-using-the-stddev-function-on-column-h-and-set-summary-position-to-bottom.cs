// Title: C# – Add StdDev Subtotals to Column H with Bottom Summary Using Aspose.Cells
// Description: Creates a workbook, populates columns A and H, defines the used range, and calls Cells.Subtotal to group by column A, calculate the standard deviation of column H, replace existing subtotals, omit page breaks, and place the summary row at the bottom. Saves the file as SubtotalStdDevBottom.xlsx.
// Keywords: Aspose.Cells C# subtotal StdDev | standard deviation subtotal .NET | bottom summary Aspose.Cells | group by column subtotal function | Cells.Subtotal example | Excel stddev subtotal code
// Common Searches: Aspose.Cells add StdDev subtotal by group C# | Cells.Subtotal bottom summary example | Calculate standard deviation subtotal in .NET | How to place subtotal summary at bottom with Aspose.Cells | C# code for StdDev subtotal on column H
// Developer Intent: Insert standard‑deviation subtotals for column H, grouped by column A, and position the summary row below the data using Aspose.Cells for .NET.
// Use Cases: Produce a sales report that groups transactions by region and shows the StdDev of revenue in column H with a final summary row. | Create a quality‑control sheet that groups test batches by type and provides StdDev subtotals for measurement values in column H, placing the summary at the end of each group. | Generate a financial analysis workbook that groups expenses by category, calculates the StdDev of amounts in column H, and adds a bottom‑positioned summary row.
// AI Prompts: Write C# code with Aspose.Cells to add StdDev subtotals for column H grouped by column A and place the summary at the bottom. | Explain each parameter of the Cells.Subtotal method when configuring a standard‑deviation subtotal with a bottom summary. | Show how to modify the example to insert page breaks between groups while still using the StdDev function for column H.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalStdDevDemo
{
    // Creates a workbook, populates columns A and H, defines the used range, and calls Cells.Subtotal to group by column A, calculate the standard deviation of column H, replace existing subtotals, omit page breaks, and place the summary row at the bottom. Saves the file as SubtotalStdDevBottom.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: fill columns A to H with some values
            // Header row
            cells["A1"].PutValue("Group");
            cells["H1"].PutValue("Values");

            // Populate data rows (example rows 2 to 11)
            for (int row = 1; row <= 10; row++)
            {
                // Group column (A) – alternate between two groups for demonstration
                cells[row, 0].PutValue(row % 2 == 0 ? "B" : "A");

                // Column H (index 7) – numeric values
                cells[row, 7].PutValue(row * 10);
            }

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;      // zero‑based index of the last used row
            int maxCol = cells.MaxDataColumn;   // zero‑based index of the last used column

            // Define the cell area that contains the data (including headers)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = maxRow,
                EndColumn = maxCol
            };

            // Apply subtotals:
            // - Group by the first column (index 0)
            // - Use StdDev function (ConsolidationFunction.StdDev)
            // - Subtotal column H (zero‑based index 7)
            // - Replace existing subtotals (true)
            // - Do not insert page breaks (false)
            // - Place summary below the data (true) → bottom position
            cells.Subtotal(
                area,
                0,                                 // group by column A
                ConsolidationFunction.StdDev,      // StdDev function
                new int[] { 7 },                   // apply to column H
                true,                              // replace existing subtotals
                false,                             // no page breaks
                true                               // summary below data (bottom)
            );

            // Save the workbook
            workbook.Save("SubtotalStdDevBottom.xlsx");
        }
    }
}
