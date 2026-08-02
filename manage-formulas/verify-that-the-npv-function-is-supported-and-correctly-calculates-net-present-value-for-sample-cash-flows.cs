// Title: Validate Aspose.Cells NPV Function in C# – Net Present Value Calculation
// Description: C# example that creates a workbook, writes cash‑flow data, applies Excel's NPV formula (including the initial investment), calculates the result, compares it with a manual .NET NPV computation, and saves the file.
// Keywords: Aspose.Cells | NPV function | C# example | net present value | Excel NPV | discount rate | formula verification | financial calculations | Aspose.Cells API | NPV verification
// Common Searches: Aspose.Cells NPV example C# | How to calculate NPV with Aspose.Cells | Verify Excel NPV result in .NET | Add initial cash flow to NPV formula Aspose.Cells | NPV tolerance check C# Aspose.Cells | Net present value calculation using Aspose.Cells
// Developer Intent: Ensure the Aspose.Cells NPV implementation returns the correct net present value for a given series of cash flows.
// Use Cases: Generate a workbook and populate cash‑flow values for financial analysis. | Apply the Excel NPV function via Aspose.Cells, adding the initial investment manually. | Programmatically compare the Aspose.Cells result with a hand‑calculated NPV to confirm accuracy. | Persist the workbook to demonstrate the full formula‑evaluation lifecycle.
// AI Prompts: Write C# code using Aspose.Cells to compute NPV for a cash‑flow array and validate the result against a manual calculation. | Show how to include the initial investment in the Excel NPV formula when using Aspose.Cells. | Create a unit test in NUnit that verifies Aspose.Cells NPV output matches expected values within a 1e‑6 tolerance. | Explain the steps to debug mismatched NPV results in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsNPVVerification
{
    // C# example that creates a workbook, writes cash‑flow data, applies Excel's NPV formula (including the initial investment), calculates the result, compares it with a manual .NET NPV computation, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data:
            // B1 = initial investment (negative cash flow)
            // B2:B5 = cash inflows for periods 1 to 4
            double rate = 0.10; // Discount rate 10%
            double[] cashFlows = { -1000, 300, 400, 500, 600 };

            // Populate the cells with the sample cash flows
            for (int i = 0; i < cashFlows.Length; i++)
            {
                // Row index i (0‑based) corresponds to Excel rows 1‑based
                cells[i, 1].PutValue(cashFlows[i]); // Column B (index 1)
            }

            // Set the NPV formula in cell C1.
            // Excel's NPV function does NOT include the initial cash flow,
            // so we add it manually: =NPV(rate, B2:B5) + B1
            string npvFormula = $"=NPV({rate},B2:B5)+B1";
            cells["C1"].Formula = npvFormula;

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the calculated NPV value
            double calculatedNpv = cells["C1"].DoubleValue;

            // Compute the expected NPV using standard .NET arithmetic for verification
            double expectedNpv = cashFlows[0]; // start with initial investment
            for (int i = 1; i < cashFlows.Length; i++)
            {
                expectedNpv += cashFlows[i] / Math.Pow(1 + rate, i);
            }

            // Output the results
            Console.WriteLine($"NPV formula: {npvFormula}");
            Console.WriteLine($"Calculated NPV (Aspose.Cells): {calculatedNpv}");
            Console.WriteLine($"Expected NPV (manual calculation): {expectedNpv}");

            // Verify that the values match within a small tolerance
            double tolerance = 1e-6;
            if (Math.Abs(calculatedNpv - expectedNpv) <= tolerance)
            {
                Console.WriteLine("Verification succeeded: NPV calculation is correct.");
            }
            else
            {
                Console.WriteLine("Verification failed: NPV calculation differs from expected value.");
            }

            // Save the workbook (optional, demonstrates the save lifecycle rule)
            workbook.Save("NPVVerification.xlsx");
        }
    }
}
