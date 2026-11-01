// Title: C# example to verify Aspose.Cells NPV function calculates net present value for a cash‑flow series
// Common Searches: how to verify NPV calculation result from Aspose.Cells in a C# program | sample C# code using Aspose.Cells to compute NPV for a series of cash flows | compare Aspose.Cells NPV function output with manual net present value calculation in C# | Aspose.Cells NPV function usage with discount rate and cash flow range example | C# tutorial for validating Excel NPV function using Aspose.Cells API
// Tags: Aspose.Cells NPV formula verification | C# Aspose.Cells financial functions example | Excel NPV calculation using Aspose.Cells API | discount rate cash flow NPV Aspose.Cells | manual NPV validation with Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook, writes a discount rate and cash‑flow series (including period‑0) to cells, applies the NPV formula =NPV(B1, B3:B6)+B2, calculates formulas, retrieves the result, computes the expected NPV manually with .NET Math.Pow, and prints both values with a verification flag confirming that Aspose.Cells' NPV matches the manual calculation.
class NpvVerification
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data: discount rate and cash flows
        double discountRate = 0.1; // 10% discount rate
        double[] cashFlows = { -1000, 300, 400, 500, 600 }; // cash flows for periods 0..4

        // Write discount rate to cell B1
        sheet.Cells["B1"].PutValue(discountRate);

        // Write cash flows to cells B2:B6 (period 0 cash flow included)
        for (int i = 0; i < cashFlows.Length; i++)
        {
            sheet.Cells[1 + i, 1].PutValue(cashFlows[i]); // Row index starts at 0, column B is index 1
        }

        // Set NPV formula in cell B8: =NPV(B1, B3:B6) + B2
        // Aspose.Cells NPV function assumes cash flows start from period 1,
        // so we separate the period‑0 cash flow (B2) and add it manually.
        sheet.Cells["B8"].Formula = "=NPV(B1, B3:B6) + B2";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the calculated NPV value
        double npvResult = sheet.Cells["B8"].DoubleValue;

        // Compute expected NPV using .NET for verification
        double expectedNpv = cashFlows[0]; // period‑0 cash flow
        for (int period = 1; period < cashFlows.Length; period++)
        {
            expectedNpv += cashFlows[period] / Math.Pow(1 + discountRate, period);
        }

        // Output the results
        Console.WriteLine($"Aspose.Cells NPV result: {npvResult}");
        Console.WriteLine($"Expected NPV (manual calculation): {expectedNpv}");
        Console.WriteLine($"Verification passed: {Math.Abs(npvResult - expectedNpv) < 1e-6}");
    }
}
