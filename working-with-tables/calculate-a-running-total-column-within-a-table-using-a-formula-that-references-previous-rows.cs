// Title: Generate a running total column in an Excel sheet with Aspose.Cells for .NET using row‑referencing formulas
// Common Searches: Aspose.Cells C# how to add a cumulative total column that references previous rows | C# set Excel formula for running total using Aspose.Cells workbook | Create running sum in Excel with Aspose.Cells .NET API | Calculate cumulative amount column programmatically with Aspose.Cells | Aspose.Cells formula for each row equals previous total plus current value
// Tags: Aspose.Cells set cell formula | cumulative sum Excel Aspose.Cells | running total column .NET | row reference formula Aspose.Cells | calculate workbook formulas Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, adds Item, Amount, and Running Total headers, fills the Amount column with sample data, assigns a formula to the first Running Total cell equal to the corresponding Amount, then sets each subsequent Running Total cell to the sum of the previous total and the current amount, evaluates all formulas, and saves the file as RunningTotal.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set headers
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Amount");
            cells["C1"].PutValue("Running Total");

            // Sample data for the Amount column (B2:B6)
            double[] amounts = { 100, 250, 150, 300, 200 };
            for (int i = 0; i < amounts.Length; i++)
            {
                cells[i + 2, 0].PutValue($"Item {i + 1}");   // Column A (Item)
                cells[i + 2, 1].PutValue(amounts[i]);       // Column B (Amount)
            }

            // Set the running total formula
            // First row (C2) just equals B2
            cells["C2"].Formula = "=B2";

            // Subsequent rows: Cn = C(n-1) + Bn (A1 style formula)
            int startRow = 3; // Excel row number (1‑based)
            int endRow = amounts.Length + 1; // because data starts at row 2
            for (int row = startRow; row <= endRow; row++)
            {
                // Build formula like "=C{row-1}+B{row}"
                string formula = $"=C{row - 1}+B{row}";
                cells[row - 1, 2].Formula = formula;
            }

            // Calculate all formulas
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("RunningTotal.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
