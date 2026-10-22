// Title: How to compute a running total column in Excel using Aspose.Cells Formula property with C#
// AI Prompts: Generate C# code that creates an Excel workbook, writes numbers to column A, and sets column B formulas to compute a running sum with Aspose.Cells. | Write a C# snippet that uses Aspose.Cells to assign a formula like =A2+B1 to each cell in column B to produce a rolling subtotal. | Provide an example that adds a header row, formats the subtotal column as currency, and evaluates all formulas using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# example for cumulative sum across rows | set formula for running total in Excel using Aspose.Cells Workbook | calculate subtotal column with previous row reference Aspose.Cells C# | how to use Formula property to create rolling total in Excel via Aspose.Cells | C# Aspose.Cells populate column B with =Arow+Bold previous row formula
// Tags: Aspose.Cells cumulative total formula C# | subtotal series using Formula property | assign previous row formula Aspose.Cells | populate column B with cumulative totals C# | calculate subtotal series Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a new workbook, writes numeric values into column A, fills column B with formulas that add each A cell to the previous B cell to produce a running total, evaluates all formulas, and saves the file as RunningTotal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        // Sample data placed in column A (A1, A2, ...)
        double[] values = { 10, 20, 15, 5, 30 };
        for (int i = 0; i < values.Length; i++)
        {
            // Put numeric value into cell A{i+1}
            sheet.Cells[i, 0].PutValue(values[i]);
        }

        // Populate column B with running total formulas
        for (int i = 0; i < values.Length; i++)
        {
            if (i == 0)
            {
                // First subtotal equals the first value
                sheet.Cells[i, 1].Formula = $"=A{i + 1}";
            }
            else
            {
                // Running total: current value + previous subtotal
                // B{i+1} = A{i+1} + B{i}
                sheet.Cells[i, 1].Formula = $"=A{i + 1}+B{i}";
            }
        }

        // Evaluate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("RunningTotal.xlsx");
    }
}
