using System;
using Aspose.Cells;

namespace AsposeCellsNPVDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: discount rate and cash flows
            double rate = 0.1; // 10% discount rate
            double[] cashFlows = { -1000, 300, 400, 500 }; // cash flow at t0, t1, t2, t3

            // Put the discount rate in cell A1
            cells["A1"].PutValue(rate);

            // Put cash flows (excluding the initial investment) in B1:B3
            // Excel's NPV function assumes the first value corresponds to period 1
            for (int i = 0; i < cashFlows.Length - 1; i++)
            {
                cells[0, i + 1].PutValue(cashFlows[i + 1]); // B1, C1, D1 ...
            }

            // Set the NPV formula in cell A2
            // Formula: =NPV(A1, B1:D1)
            cells["A2"].Formula = "=NPV(A1, B1:D1)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the calculated NPV value
            double npvCalculated = cells["A2"].DoubleValue;

            // Compute expected NPV using .NET for verification
            double npvExpected = 0.0;
            for (int i = 1; i < cashFlows.Length; i++)
            {
                npvExpected += cashFlows[i] / Math.Pow(1 + rate, i);
            }

            // Output results
            Console.WriteLine($"Discount Rate (A1): {rate}");
            Console.WriteLine("Cash Flows (t1..t3):");
            for (int i = 1; i < cashFlows.Length; i++)
            {
                Console.WriteLine($"  Period {i}: {cashFlows[i]}");
            }
            Console.WriteLine($"NPV calculated by Aspose.Cells (A2): {npvCalculated}");
            Console.WriteLine($"NPV expected (manual calculation): {npvExpected}");
            Console.WriteLine($"Difference: {Math.Abs(npvCalculated - npvExpected)}");

            // Save the workbook (optional)
            workbook.Save("NPVDemo.xlsx");
        }
    }
}