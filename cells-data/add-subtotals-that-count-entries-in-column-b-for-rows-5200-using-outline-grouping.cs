// Title: C# – Add outline subtotals that count column B (rows 5‑200) with Aspose.Cells for .NET
// Description: This example creates a workbook, optionally fills rows 5‑200 with group keys in column A and numeric values in column B, defines a CellArea for that range, and uses the Cells.Subtotal method with ConsolidationFunction.Count to insert count subtotals grouped by column B. The outline’s SummaryRowBelow property is set so each subtotal appears below its group, and the file is saved as SubtotalOutlineDemo.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | .NET | subtotal count | outline grouping | Excel subtotal | CellArea | ConsolidationFunction.Count | summary row below | automated Excel report
// Common Searches: Aspose.Cells count subtotal column B C# | outline grouping subtotal rows 5 to 200 Aspose.Cells | how to add subtotal with summary row below using Aspose.Cells | C# Subtotal method Aspose.Cells example | generate Excel subtotals programmatically .NET
// Developer Intent: Insert count subtotals for column B (rows 5‑200) using outline grouping and place the summary rows below each detail group.
// Use Cases: Create a categorized report where each category shows the number of items it contains. | Automate Excel generation for large data sets without manual subtotal entry. | Produce a printable worksheet with outline‑style totals for quick data analysis.
// AI Prompts: Generate C# code with Aspose.Cells that adds a count subtotal for column B (rows 5‑200) and sets SummaryRowBelow to true. | Explain the parameters of the Cells.Subtotal method for grouping, function selection, and result column. | Show how to customize outline formatting (e.g., collapse/expand) after adding subtotals with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // This example creates a workbook, optionally fills rows 5‑200 with group keys in column A and numeric values in column B, defines a CellArea for that range, and uses the Cells.Subtotal method with ConsolidationFunction.Count to insert count subtotals grouped by column B. The outline’s SummaryRowBelow property is set so each subtotal appears below its group, and the file is saved as SubtotalOutlineDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // (Optional) Populate sample data in columns A and B for rows 5‑200
            // Here we just fill column B with some values; column A can hold group keys
            for (int row = 4; row <= 199; row++) // zero‑based index: row 5 = 4
            {
                // Example group key in column A (could be any string)
                cells[row, 0].PutValue("Group" + ((row - 4) % 5 + 1));
                // Value in column B
                cells[row, 1].PutValue(row);
            }

            // Define the cell area that includes columns A and B, rows 5‑200
            CellArea area = new CellArea
            {
                StartRow = 4,      // row 5 (zero‑based)
                StartColumn = 0,   // column A
                EndRow = 199,      // row 200 (zero‑based)
                EndColumn = 1      // column B
            };

            // Add subtotals:
            // - Group by column B (index 1)
            // - Use Count function to count entries
            // - Apply the subtotal to column B (index 1)
            cells.Subtotal(area, 1, ConsolidationFunction.Count, new int[] { 1 });

            // Ensure the summary row appears below the detail rows in the outline
            worksheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("SubtotalOutlineDemo.xlsx");
        }
    }
}
