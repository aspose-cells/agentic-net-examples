// Title: Validate Aspose.Cells NPV Function in C# – Net Present Value Calculation
// Description: C# sample that creates a workbook, writes a discount rate and cash‑flow series, applies the Excel =NPV formula, evaluates it via Worksheet.CalculateFormula and workbook.CalculateFormula, compares the result with a manual Math.Pow computation, and confirms the Aspose.Cells NPV implementation is supported and accurate.
// Keywords: Aspose.Cells NPV | C# net present value | Excel NPV formula evaluation | Aspose.Cells formula API | calculate NPV with Aspose | financial functions .NET | verify Excel functions programmatically
// Common Searches: how to use npv function in aspose.cells | aspnet verify npv calculation | c# aspose.cells net present value example | compare manual npv with aspose.cells result | excel npv formula in aspose.cells workbook
// Developer Intent: Ensure the NPV function exists in Aspose.Cells and returns the correct net present value for a given discount rate and cash‑flow array.
// Use Cases: Populate discount rate and cash‑flow values in cells, then assign =NPV to compute financial metrics. | Retrieve the NPV instantly with Worksheet.CalculateFormula(string) for on‑the‑fly validation. | Run workbook.CalculateFormula() to update the worksheet, read the cell value, and assert it matches a manually calculated reference.
// AI Prompts: Write C# code using Aspose.Cells that calculates NPV for a cash‑flow list and verifies the result against a manual formula. | Explain the argument handling of the Excel NPV function inside Aspose.Cells and how to extract the computed value programmatically. | Create an MSTest unit test that asserts Aspose.Cells NPV output equals the expected value within a 1e-6 tolerance.

using System;
using Aspose.Cells;

namespace AsposeCellsNPVVerification
{
    // C# sample that creates a workbook, writes a discount rate and cash‑flow series, applies the Excel =NPV formula, evaluates it via Worksheet.CalculateFormula and workbook.CalculateFormula, compares the result with a manual Math.Pow computation, and confirms the Aspose.Cells NPV implementation is supported and accurate.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data:
            // A1 – discount rate
            // A2:A4 – cash flow values (period 1 to 3)
            // A5 – formula to calculate NPV
            double discountRate = 0.10; // 10%
            double[] cashFlows = { 300, 420, 680 };

            // Populate the worksheet
            cells["A1"].PutValue(discountRate);
            for (int i = 0; i < cashFlows.Length; i++)
            {
                // Cash flows start at row 2 (index 1)
                cells[i + 1, 0].PutValue(cashFlows[i]);
            }

            // Set the NPV formula in cell A5.
            // Note: Excel's NPV function assumes the first cash flow occurs at the end of period 1,
            // so we only pass the range A2:A4.
            cells["A5"].Formula = "=NPV(A1, A2:A4)";

            // Calculate the formula using the worksheet's CalculateFormula method.
            // This returns the evaluated result directly.
            object npvResult = sheet.CalculateFormula("=NPV(A1, A2:A4)");

            // Also calculate all formulas in the workbook to ensure the cell value is updated.
            workbook.CalculateFormula();

            // Retrieve the value stored in cell A5 after calculation.
            double npvFromCell = cells["A5"].DoubleValue;

            // Expected NPV calculated manually for verification.
            double expectedNpv = cashFlows[0] / Math.Pow(1 + discountRate, 1) +
                                 cashFlows[1] / Math.Pow(1 + discountRate, 2) +
                                 cashFlows[2] / Math.Pow(1 + discountRate, 3);

            // Output the results.
            Console.WriteLine("Discount Rate (A1): " + discountRate);
            Console.WriteLine("Cash Flows (A2:A4): " + string.Join(", ", cashFlows));
            Console.WriteLine("NPV calculated via CalculateFormula(string): " + npvResult);
            Console.WriteLine("NPV value stored in cell A5 after workbook.CalculateFormula(): " + npvFromCell);
            Console.WriteLine("Expected NPV (manual calculation): " + expectedNpv);
            Console.WriteLine("Verification: " + (Math.Abs(Convert.ToDouble(npvResult) - expectedNpv) < 1e-6
                                                    ? "PASS"
                                                    : "FAIL"));

            // Save the workbook (optional, demonstrates the save rule)
            workbook.Save("NPVVerification.xlsx");
        }
    }
}
