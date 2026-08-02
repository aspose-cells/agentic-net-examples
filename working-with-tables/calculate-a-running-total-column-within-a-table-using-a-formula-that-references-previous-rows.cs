// Title: C# – Add a Running Total column to an Aspose.Cells ListObject using row‑reference formulas
// Description: Creates a workbook, defines an Amount column, converts A1:B6 into a ListObject, shows the totals row, and fills the RunningTotal column with formulas that reference the previous row. The workbook is calculated and saved as RunningTotalTable.xlsx.
// Keywords: Aspose.Cells C# ListObject running total | Excel table cumulative sum formula | Aspose.Cells calculate running total | C# add column with previous row reference | structured table formula Aspose.Cells | progressive sum in Excel using Aspose
// Common Searches: Aspose.Cells add cumulative column to table C# | how to reference previous row in Aspose.Cells formula | create running total in Excel ListObject with Aspose | C# example for progressive sum in Aspose.Cells table | populate table column with formula Aspose.Cells
// Developer Intent: Generate a RunningTotal column inside an Excel ListObject that computes a cumulative sum by referencing the prior row’s total using Aspose.Cells in C#.
// Use Cases: Invoice worksheets that display a running subtotal for each line item. | Daily sales dashboards where each new entry updates the progressive total automatically. | Budget trackers that recalculate cumulative expenses whenever the Amount column changes.
// AI Prompts: Write C# code with Aspose.Cells to add a RunningTotal column to an existing ListObject, using formulas that reference the previous row. | Show how to modify the example so the running total updates automatically when rows are inserted or deleted in the table. | Provide a version that uses structured table references (e.g., [@Amount]) instead of absolute cell addresses for the running total formula.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, defines an Amount column, converts A1:B6 into a ListObject, shows the totals row, and fills the RunningTotal column with formulas that reference the previous row. The workbook is calculated and saved as RunningTotalTable.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ----- 1. Prepare sample data -----
        // Header row
        cells["A1"].PutValue("Amount");
        cells["B1"].PutValue("RunningTotal");

        // Sample amounts (rows 2‑6)
        double[] amounts = { 100, 250, 175, 300, 225 };
        for (int i = 0; i < amounts.Length; i++)
        {
            cells[i + 1, 0].PutValue(amounts[i]); // Column A
        }

        // ----- 2. Create a table that includes both columns -----
        // Table range: A1:B6 (header + 5 data rows)
        int tableIndex = sheet.ListObjects.Add(0, 0, amounts.Length, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Show the totals row (optional)
        table.ShowTotals = true;

        // ----- 3. Populate the RunningTotal column with a formula that references the previous row -----
        // The RunningTotal column is the second column in the table (column offset = 1)
        for (int i = 0; i < amounts.Length; i++)
        {
            int rowOffset = i + 1;          // Row offset inside the table (skip header)
            int colOffset = 1;              // RunningTotal column offset

            string formula;
            if (i == 0)
            {
                // First row: running total equals the amount itself
                // A2 is the first amount cell (row 2 in the worksheet)
                formula = $"=A{rowOffset + 1}";
            }
            else
            {
                // Subsequent rows: current amount + previous running total
                // A{row} is current amount, B{row-1} is previous total
                int currentRow = rowOffset + 1;      // Worksheet row number for current amount
                int previousRow = currentRow - 1;    // Worksheet row number for previous total
                formula = $"=A{currentRow}+B{previousRow}";
            }

            // Apply the formula to the cell inside the table
            table.PutCellFormula(rowOffset, colOffset, formula);
        }

        // ----- 4. Calculate all formulas -----
        workbook.CalculateFormula();

        // ----- 5. Save the workbook -----
        workbook.Save("RunningTotalTable.xlsx");
    }
}
