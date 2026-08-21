// Title: Aspose.Cells .NET – Apply Subtotal to Full Data Range Using MaxDataRow
// Description: Creates a workbook, fills it with a header and sample rows, finds the last populated row with MaxDataRow, defines a CellArea that spans the entire dataset, and calls Cells.Subtotal to group by the first column, sum the Amount column, replace existing subtotals, insert page breaks, and place summary rows below the data.
// Keywords: Aspose.Cells Subtotal .NET | Cells.Subtotal dynamic range | MaxDataRow last row | group by column subtotal | Excel subtotal programmatically | page breaks with subtotal | summary rows below data
// Common Searches: Aspose.Cells how to subtotal entire range | use MaxDataRow with Cells.Subtotal | add subtotal rows programmatically in .NET | group data and sum column Aspose.Cells | insert page breaks when applying subtotal
// Developer Intent: Add subtotal rows to a worksheet covering all populated rows without hard‑coding the end row.
// Use Cases: Generate a printable sales report that groups items by category and shows total amounts. | Create an Excel export that automatically adapts to any dataset size and adds subtotals. | Produce a multi‑page financial statement with page breaks and summary rows inserted after each group.
// AI Prompts: Show how to add average and count subtotals for the Amount column while still using MaxDataRow. | Demonstrate using Cells.Subtotal with multiple summary columns (e.g., sum, min, max) in a single call.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // Creates a workbook, fills it with a header and sample rows, finds the last populated row with MaxDataRow, defines a CellArea that spans the entire dataset, and calls Cells.Subtotal to group by the first column, sum the Amount column, replace existing subtotals, insert page breaks, and place summary rows below the data.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including a header row)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Amount");

            object[,] data = new object[,]
            {
                { "Food", "Apple", 120 },
                { "Food", "Bread", 80 },
                { "Drink", "Water", 50 },
                { "Drink", "Juice", 150 },
                { "Snack", "Chips", 70 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
                for (int c = 0; c < data.GetLength(1); c++)
                    cells[r + 1, c].PutValue(data[r, c]);

            // Determine the last row that contains data
            int lastDataRow = cells.MaxDataRow; // zero‑based index

            // Define the cell area covering the entire data range (including header)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = lastDataRow,
                EndColumn = 2 // three columns: A, B, C
            };

            // Apply subtotal:
            // - Group by the first column (Category) -> index 0
            // - Use SUM function
            // - Add subtotal for the third column (Amount) -> index 2
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, true, true);

            // Save the workbook
            workbook.Save("SubtotalFullRange.xlsx");
        }
    }
}
