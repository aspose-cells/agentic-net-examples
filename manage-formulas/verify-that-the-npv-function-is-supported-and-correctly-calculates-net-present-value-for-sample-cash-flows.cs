// Title: Validate Aspose.Cells NPV formula computes correct net present value for a cash‑flow series in C#
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, inserts a discount rate and a series of cash‑flow values, applies the Excel NPV formula to the cash‑flow range (adding the period‑0 cash flow separately), and compares the result with a manually computed NPV. | Write a C# verification routine using Aspose.Cells that asserts the NPV function output matches the expected value within a tolerance of 1e‑6.
// Common Searches: aspacells c# calculate NPV using Excel formula | verify NPV function result against manual calculation with Aspose.Cells | c# example Aspose.Cells net present value of cash flow series | Aspose.Cells NPV formula tolerance check in C# program
// Tags: Aspose.Cells NPV function verification | populate discount rate and cash flows in workbook cells | compare Aspose.Cells NPV output with manual calculation | assert NPV result within tolerance in C# | Excel NPV formula integration via Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a workbook, fills cells with a discount rate and cash‑flow series, uses the Excel NPV formula through Aspose.Cells (adding the period‑0 cash flow separately), calculates the sheet, manually computes the expected NPV, and confirms the two values match within a 1e‑6 tolerance.
class NpvVerification
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data:
        // Discount rate in B1
        // Initial cash flow (period 0) in B2
        // Cash flows for periods 1..3 in B3:B5
        double discountRate = 0.10; // 10%
        double[] cashFlows = { -1000, 300, 400, 500 };

        // Populate cells
        cells["B1"].PutValue(discountRate);
        cells["B2"].PutValue(cashFlows[0]); // period 0
        cells["B3"].PutValue(cashFlows[1]); // period 1
        cells["B4"].PutValue(cashFlows[2]); // period 2
        cells["B5"].PutValue(cashFlows[3]); // period 3

        // Use Excel's NPV function (which discounts cash flows starting from period 1)
        // Add the period‑0 cash flow separately.
        // Formula in B6: =NPV(B1, B3:B5) + B2
        cells["B6"].Formula = "=NPV(B1, B3:B5) + B2";

        // Calculate all formulas
        workbook.CalculateFormula();

        // Retrieve the calculated NPV
        double calculatedNpv = cells["B6"].DoubleValue;

        // Compute expected NPV manually for verification
        double expectedNpv = cashFlows[0];
        for (int i = 1; i < cashFlows.Length; i++)
        {
            expectedNpv += cashFlows[i] / Math.Pow(1 + discountRate, i);
        }

        // Compare with a tolerance
        double tolerance = 1e-6;
        bool isCorrect = Math.Abs(calculatedNpv - expectedNpv) < tolerance;

        // Output results
        Console.WriteLine($"Calculated NPV (Aspose.Cells): {calculatedNpv}");
        Console.WriteLine($"Expected NPV (manual): {expectedNpv}");
        Console.WriteLine($"NPV function verification: {(isCorrect ? "PASS" : "FAIL")}");

        // Optional: save workbook for visual inspection
        // workbook.Save("NPVVerification.xlsx");
    }
}
