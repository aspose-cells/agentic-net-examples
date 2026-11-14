// Title: Calculate a running total column in Excel with Aspose.Cells Formula property using C#
// Common Searches: C# Aspose.Cells example for cumulative sum column in Excel | How to set a running total formula programmatically with Aspose.Cells | Using Aspose.Cells Formula to add current cell value to previous total in C# | Create dynamic Excel running total using Aspose.Cells workbook calculation | Aspose.Cells calculate running total across rows C# code sample
// Tags: Aspose.Cells running total formula C# | dynamic Excel formula assignment Aspose.Cells | cumulative sum column generation C# | programmatic workbook calculation Aspose.Cells | Excel running total automation C#

using System;
using Aspose.Cells;

// The sample creates a new workbook, fills column A with numeric values, adds a "Running Total" header in column B, assigns B2 to reference A2, then dynamically builds formulas for B3‑B5 that add each A‑column value to the previous B‑column subtotal, evaluates all formulas, and saves the result as RunningTotal.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook (using the provided create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column A with sample values (e.g., 10, 20, 30, 40)
        sheet.Cells["A1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].PutValue(20);
        sheet.Cells["A4"].PutValue(30);
        sheet.Cells["A5"].PutValue(40);

        // Add header for running total column
        sheet.Cells["B1"].PutValue("Running Total");

        // Set formula for the first running total cell (B2 = A2)
        sheet.Cells["B2"].Formula = "A2";

        // Set formulas for subsequent cells to add current value to previous subtotal
        // B3 = B2 + A3, B4 = B3 + A4, etc.
        for (int row = 3; row <= 5; row++)
        {
            // Build formula string dynamically
            string formula = $"B{row - 1}+A{row}";
            sheet.Cells[$"B{row}"].Formula = formula;
        }

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook (using the provided save rule)
        workbook.Save("RunningTotal.xlsx");
    }
}
