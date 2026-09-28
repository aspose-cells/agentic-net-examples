// Title: Create grouped SUM subtotals for column D (rows 2‑100) in an Excel workbook using Aspose.Cells C#
// AI Prompts: Write C# code that defines a CellArea covering rows 2‑100 and columns A‑D, then calls Cells.Subtotal to group by column A and insert SUM subtotals for column D below each group. | Show how to modify an existing Aspose.Cells workbook to add subtotal rows for the Amount column based on categories in column A, using ConsolidationFunction.Sum. | Generate a complete C# example that creates a workbook, populates sample data, applies grouped subtotals with Cells.Subtotal, and saves the file as SubtotalResult.xlsx.
// Common Searches: aspnet c# add subtotal rows per category with Aspose.Cells | using Aspose.Cells Subtotal method to sum values in column D for rows 2 to 100 | group by column A and calculate sum subtotals in Excel via Aspose.Cells C# example | Aspose.Cells create subtotal rows below each group in a worksheet
// Tags: Aspose.Cells Cells.Subtotal method | C# generate Excel subtotals | grouped sum subtotal column D | define CellArea for Excel range | save workbook as .xlsx with Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    // Demonstrates creating a workbook, filling rows 2‑100 with sample data, defining a CellArea for A‑D, and using Cells.Subtotal to group by column A and insert SUM subtotals for column D at each group's bottom, then saving the result as SubtotalResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data for rows 2‑100 (zero‑based rows 1‑99)
            // Columns: A (Category), B (Item), C (Description), D (Amount)
            for (int row = 1; row <= 99; row++)
            {
                // Simple grouping: alternate between two categories in column A
                string category = (row % 2 == 0) ? "Group1" : "Group2";
                cells[row, 0].PutValue(category);          // Column A
                cells[row, 1].PutValue($"Item{row}");      // Column B
                cells[row, 2].PutValue($"Desc{row}");      // Column C
                cells[row, 3].PutValue(row * 10);          // Column D (numeric value to sum)
            }

            // Define the cell area covering rows 2‑100 and columns A‑D
            CellArea area = new CellArea
            {
                StartRow = 1,      // Row 2 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 99,       // Row 100
                EndColumn = 3      // Column D
            };

            // Apply subtotals:
            // - Group by column A (index 0)
            // - Use SUM function
            // - Subtotal column D (index 3)
            // - Replace existing subtotals, no page breaks, summary placed below each group
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 3 }, true, false, true);

            // Save the workbook
            workbook.Save("SubtotalResult.xlsx");
        }
    }
}
