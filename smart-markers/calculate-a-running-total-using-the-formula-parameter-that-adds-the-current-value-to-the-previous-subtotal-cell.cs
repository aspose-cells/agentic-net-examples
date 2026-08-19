// Title: C# – Compute a Running Total Column in Excel Using Aspose.Cells Formula
// Description: Demonstrates how to create a workbook, add Item, Amount, and Running Total headers, populate sample data, set the first total cell to the first amount, then assign a Formula to each subsequent cell that adds the current Amount to the previous Running Total, recalculate all formulas, and save the file as RunningTotalDemo.xlsx. The loop works for any number of rows.
// Keywords: Aspose.Cells running total | C# Excel cumulative sum | Formula property Aspose.Cells | calculate running total .NET | Excel subtotal column programmatically | Aspose.Cells CalculateFormula | dynamic running balance C# | smart markers cumulative total | Excel automation Aspose | Workbook.Save running total
// Common Searches: Aspose.Cells set formula for running total | C# cumulative sum column Excel | How to calculate running total with Aspose.Cells | Formula property example Aspose.Cells .NET | Create running balance worksheet using Aspose
// Developer Intent: Add a running‑total column by programmatically assigning a Formula that adds the current Amount cell to the previous row’s total and then recalculate the workbook.
// Use Cases: Generate an invoice sheet that shows a cumulative amount paid per line item. | Build a bank‑statement style report where each transaction updates the running balance automatically. | Create a sales dashboard that reflects a live cumulative total as new sales figures are entered.
// AI Prompts: Write C# code with Aspose.Cells to add a running‑total column that references the previous row’s total cell. | Show how to use the Formula property in a loop to compute a cumulative sum and then call CalculateFormula. | Explain how to adapt the formula loop for an unknown number of data rows and avoid off‑by‑one errors.

using Aspose.Cells;
using System;

// Demonstrates how to create a workbook, add Item, Amount, and Running Total headers, populate sample data, set the first total cell to the first amount, then assign a Formula to each subsequent cell that adds the current Amount to the previous Running Total, recalculate all formulas, and save the file as RunningTotalDemo.xlsx. The loop works for any number of rows.
class RunningTotalDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add headers
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Amount");
        cells["C1"].PutValue("Running Total");

        // Sample data
        string[] items = { "A", "B", "C", "D" };
        double[] amounts = { 100, 250, 150, 300 };

        // Populate data rows
        for (int i = 0; i < items.Length; i++)
        {
            cells[i + 1, 0].PutValue(items[i]);   // Column A
            cells[i + 1, 1].PutValue(amounts[i]); // Column B
        }

        // First running total equals the first amount
        cells[1, 2].Formula = "=B2";

        // Set running total formula for the rest of the rows:
        // RunningTotal(row) = RunningTotal(previous row) + Amount(current row)
        for (int row = 2; row <= items.Length; row++)
        {
            // Excel row numbers are 1‑based, so add 1 to the zero‑based index
            string formula = $"=C{row}+B{row + 1}";
            cells[row, 2].Formula = formula;
        }

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the result
        workbook.Save("RunningTotalDemo.xlsx");
    }
}
