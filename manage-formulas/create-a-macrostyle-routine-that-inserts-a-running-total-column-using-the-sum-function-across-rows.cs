// Title: C# macro‑style routine to add a running total column with SUM formulas using Aspose.Cells
// Description: Shows how to create a workbook, fill columns A and B with sample data, add a "RunningTotal" header in column C, and programmatically place a cumulative SUM formula (=SUM($B$2:B{row})) in each row before saving as RunningTotalDemo.xlsx.
// Keywords: Aspose.Cells | .NET | C# | running total column | SUM formula | macro style routine | cumulative total | Excel automation | insert formula programmatically | worksheet calculations
// Common Searches: Aspose.Cells add running total column C# | How to insert cumulative SUM formula with Aspose.Cells .NET | Macro‑style example for running totals in Excel using Aspose.Cells | Programmatically create running total column in C# workbook | Aspose.Cells formula insertion per row
// Developer Intent: Add a column that automatically calculates a cumulative total for each row using the SUM function, implemented in C# with Aspose.Cells.
// Use Cases: Financial statements that display a progressive expense total per line item. | Sales dashboards where each row shows the accumulated revenue up to that point. | Invoice templates that automatically compute a running balance as new items are entered.
// AI Prompts: Generate a C# method using Aspose.Cells that adds a cumulative total column to any worksheet, handling an arbitrary number of rows. | Provide a macro‑style Aspose.Cells snippet that updates the running total column when additional rows are appended. | Explain how to adjust the running total formula to start from a different column or to incorporate conditional summing in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsMacroStyle
{
    // Shows how to create a workbook, fill columns A and B with sample data, add a "RunningTotal" header in column C, and programmatically place a cumulative SUM formula (=SUM($B$2:B{row})) in each row before saving as RunningTotalDemo.xlsx.
    public class RunningTotalRoutine
    {
        public static void InsertRunningTotal()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: Header in A1 and B1
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Amount");

            // Populate sample rows (A2:B6)
            string[] items = { "Item1", "Item2", "Item3", "Item4", "Item5" };
            double[] amounts = { 100, 150, 200, 250, 300 };
            for (int i = 0; i < items.Length; i++)
            {
                cells[i + 1, 0].PutValue(items[i]);   // Column A
                cells[i + 1, 1].PutValue(amounts[i]); // Column B
            }

            // Add header for Running Total column in C1
            cells["C1"].PutValue("RunningTotal");

            // Insert running total formula for each data row
            // Formula: =SUM($B$2:B{currentRow})
            for (int row = 1; row <= items.Length; row++) // row index is zero‑based
            {
                // Build the address for the end of the range (e.g., B2, B3, ...)
                string endAddress = CellsHelper.CellIndexToName(row, 1); // column B = index 1
                string formula = $"=SUM($B$2:{endAddress})";
                cells[row, 2].Formula = formula; // Column C = index 2
            }

            // Save the workbook (save rule)
            workbook.Save("RunningTotalDemo.xlsx");
        }

        // Entry point for demonstration
        public static void Main()
        {
            InsertRunningTotal();
            Console.WriteLine("Workbook with running total column created successfully.");
        }
    }
}
