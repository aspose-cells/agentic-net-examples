// Title: Insert a running‑total formula that references the previous row using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add a cumulative sum column that uses the previous row in the formula | how to programmatically create a running total in Excel with Aspose.Cells .NET | insert dynamic formula referencing prior row for running totals using Aspose.Cells | C# example for calculating running totals with Excel formulas via Aspose.Cells | generate running total column in workbook with Aspose.Cells and formula recursion
// Tags: insert running total formula Aspose.Cells | previous row reference Excel formula C# | cumulative sum column programmatically .NET | dynamic formula insertion Aspose.Cells | calculate running totals workbook Aspose.Cells | Excel running total using cell formulas C#

using Aspose.Cells;
using System;

// The sample creates a new workbook, writes numeric values to column A, then adds formulas to column B that compute a running total by referencing the previous row's total (first row uses =A2, subsequent rows use =B{prev}+A{current}). After recalculating formulas, the workbook is saved as RunningTotal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data to calculate running totals for (placed in column A, starting at row 2)
        double[] values = { 10, 20, 30, 40, 50 };
        for (int i = 0; i < values.Length; i++)
        {
            // Row index is zero‑based; Excel row 2 corresponds to index 1
            sheet.Cells[i + 1, 0].PutValue(values[i]);
        }

        // Insert running total formula in column B.
        // First data row: total equals its own value.
        // Subsequent rows: total = previous total + current value.
        for (int i = 0; i < values.Length; i++)
        {
            int excelRow = i + 2; // Excel row number (1‑based, header row assumed)
            if (i == 0)
            {
                // First row formula: =A2
                sheet.Cells[excelRow - 1, 1].Formula = $"=A{excelRow}";
            }
            else
            {
                // Subsequent rows formula: =B{prevRow}+A{currentRow}
                sheet.Cells[excelRow - 1, 1].Formula = $"=B{excelRow - 1}+A{excelRow}";
            }
        }

        // Recalculate all formulas so the workbook contains the computed values
        workbook.CalculateFormula();

        // Save the workbook (lifecycle save rule)
        workbook.Save("RunningTotal.xlsx");
    }
}
