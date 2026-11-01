// Title: Validate that MINIFS respects AutoFilter visibility when calculating formulas with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# calculate MINIFS after applying AutoFilter | How to ensure MINIFS ignores hidden rows in a filtered worksheet using Aspose.Cells | C# example of using Workbook.CalculateFormula with AutoFilter applied | Validate MINIFS result with visible rows only in Aspose.Cells .NET
// Tags: MINIFS with AutoFilter in Aspose.Cells | Workbook.CalculateFormula respecting worksheet filter | filtered rows calculation in Aspose.Cells C# | AutoFilter visibility handling for formulas | C# Aspose.Cells MINIFS filtered data

using Aspose.Cells;
using System;

// The example creates a workbook, fills columns A and B with numeric values and categories, applies an AutoFilter on column B to show only rows where the category is "X", inserts a MINIFS formula in C2 that should evaluate only the visible rows, forces formula evaluation with Workbook.CalculateFormula, verifies the result equals the expected minimum (8), prints pass/fail, and saves the workbook as MinIfsFilterTest.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate sample data
        ws.Cells["A1"].PutValue("Value");
        ws.Cells["B1"].PutValue("Category");
        ws.Cells["A2"].PutValue(10);
        ws.Cells["B2"].PutValue("X");
        ws.Cells["A3"].PutValue(5);
        ws.Cells["B3"].PutValue("Y");
        ws.Cells["A4"].PutValue(8);
        ws.Cells["B4"].PutValue("X");
        ws.Cells["A5"].PutValue(3);
        ws.Cells["B5"].PutValue("Y");

        // Apply an AutoFilter to column B and filter for "X"
        ws.AutoFilter.Range = "A1:B5";
        ws.AutoFilter.AddFilter(1, "X"); // Column index 1 corresponds to column B
        ws.AutoFilter.Refresh();

        // Insert MINIFS formula that should consider only visible (filtered) rows
        ws.Cells["C2"].Formula = "MINIFS(A2:A5,B2:B5,\"X\")";

        // Force calculation of all formulas
        wb.CalculateFormula();

        // Retrieve and display the result
        double result = ws.Cells["C2"].DoubleValue;
        Console.WriteLine("MINIFS result with filter applied: " + result);

        // Expected minimum for category "X" is 8 (values 10 and 8)
        if (Math.Abs(result - 8) < 0.0001)
            Console.WriteLine("PASS: Filter respected.");
        else
            Console.WriteLine("FAIL: Filter not respected.");

        // Save the workbook for inspection (optional)
        wb.Save("MinIfsFilterTest.xlsx");
    }
}
