// Title: Create a running total column in Excel with Aspose.Cells C# by setting cell formulas programmatically
// AI Prompts: Write C# code with Aspose.Cells that fills column A with numbers and sets column B formulas to compute an incremental total across rows. | Show how to assign the first subtotal cell directly and then generate relative formulas for the remaining rows that add the previous subtotal and the current A‑cell value. | Demonstrate calling workbook.CalculateFormula and persisting the workbook after applying incremental‑total formulas using Aspose.Cells for .NET.
// Common Searches: aspnet how to calculate cumulative sum in Excel using Aspose.Cells formula property | c# Aspose.Cells set subtotal formula for each row in a worksheet | example of using Formula property to create a total column with Aspose.Cells | programmatically generate cumulative totals in Excel file with Aspose.Cells .NET
// Tags: Aspose.Cells set cell formula C# | running total column Excel Aspose.Cells | formula‑based total computation Aspose.Cells | evaluate workbook formulas Aspose.Cells | export workbook to xlsx Aspose.Cells

using System;
using Aspose.Cells;

// // This example creates a new workbook, writes a series of numeric values into column A, assigns formulas to column B that compute a running total (the first cell copies A2, each subsequent cell adds the previous subtotal in column B to the current value in column A), forces formula evaluation with CalculateFormula, and saves the file as RunningTotal.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data placed in column A (A2:A6)
        double[] values = { 100, 200, 150, 250, 300 };
        for (int i = 0; i < values.Length; i++)
        {
            // Row index in Aspose.Cells is zero‑based, so i+1 corresponds to Excel row 2,3,...
            cells[i + 1, 0].PutValue(values[i]); // Column A
        }

        // Set the running total formula in column B
        // First subtotal just copies the first value
        cells[1, 1].Formula = "=A2"; // B2 = A2

        // Subsequent rows add the current value to the previous subtotal
        // Excel rows start at 1, so we work with Excel row numbers for the formula string
        for (int excelRow = 3; excelRow <= values.Length + 1; excelRow++)
        {
            // B{excelRow} = B{excelRow-1} + A{excelRow}
            string formula = $"=B{excelRow - 1}+A{excelRow}";
            // Convert Excel row back to zero‑based index for the Cells collection
            cells[excelRow - 1, 1].Formula = formula;
        }

        // Calculate all formulas so the running totals are materialized
        workbook.CalculateFormula();

        // Save the workbook (uses the provided save rule)
        workbook.Save("RunningTotal.xlsx");
    }
}
