// Title: How to verify Aspose.Cells for .NET calculates the NPV function correctly for a cash flow series in C#
// Common Searches: how to test NPV support in Aspose.Cells using C# | sample code to calculate net present value with Aspose.Cells workbook | validate Excel NPV formula results against manual calculation in .NET
// Tags: Aspose.Cells NPV functionality check | C# workbook net present value demo | Excel NPV formula via Aspose.Cells API | discount rate cash flow evaluation .NET | precision tolerance comparison C#

using System;
using Aspose.Cells;

// The example creates a workbook, sets a 10% discount rate and a series of cash flows, applies the Excel NPV formula (adding the initial investment separately) via Aspose.Cells, recalculates formulas, and compares the result with a manually computed NPV using a 1e-6 tolerance, confirming whether the NPV function is supported and accurate.
class NpvVerification
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data: discount rate and cash flows
        // A1: discount rate (10%)
        cells["A1"].PutValue(0.10);
        // B1: initial investment (cash outflow)
        cells["B1"].PutValue(-1000.0);
        // B2:B4: cash inflows for periods 1, 2, 3
        cells["B2"].PutValue(300.0);
        cells["B3"].PutValue(400.0);
        cells["B4"].PutValue(500.0);

        // Use Excel's NPV function.
        // NPV(rate, values) calculates the present value of cash flows *after* period 0.
        // Therefore we add the initial investment (B1) separately.
        // C1 will contain the calculated NPV.
        cells["C1"].Formula = "=NPV(A1,B2:B4)+B1";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the calculated NPV value
        double calculatedNpv = cells["C1"].DoubleValue;

        // Expected NPV calculated manually:
        // NPV = -1000 + 300/1.1 + 400/1.1^2 + 500/1.1^3
        double expectedNpv = -1000.0
                             + 300.0 / Math.Pow(1.10, 1)
                             + 400.0 / Math.Pow(1.10, 2)
                             + 500.0 / Math.Pow(1.10, 3);

        // Define a tolerance for floating‑point comparison
        double tolerance = 1e-6;

        // Verify the result
        if (Math.Abs(calculatedNpv - expectedNpv) <= tolerance)
        {
            Console.WriteLine("NPV function is supported and calculation is correct.");
            Console.WriteLine($"Calculated NPV: {calculatedNpv:F6}");
            Console.WriteLine($"Expected NPV:   {expectedNpv:F6}");
        }
        else
        {
            Console.WriteLine("NPV calculation mismatch.");
            Console.WriteLine($"Calculated NPV: {calculatedNpv:F6}");
            Console.WriteLine($"Expected NPV:   {expectedNpv:F6}");
        }

        // (Optional) Save the workbook for inspection
        // workbook.Save("NPV_Verification.xlsx");
    }
}
