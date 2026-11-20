// Title: Validate MINIFS formula results in Excel 2016 compatibility mode using Aspose.Cells for .NET (C#)
// Common Searches: how to calculate MINIFS with Aspose.Cells in C# when workbook is set to Excel 2016 compatibility | C# code to verify MINIFS function returns correct value using Aspose.Cells | Aspose.Cells example for testing MINIFS formula in Excel 2016 mode | programmatically validate MINIFS results in a .NET workbook with compatibility settings | unit test MINIFS calculation with Aspose.Cells and Excel 2016 compatibility flag
// Tags: Aspose.Cells MINIFS formula calculation | Excel 2016 compatibility mode workbook | C# validate MINIFS result | formula evaluation Aspose.Cells .NET | programmatic MINIFS testing

using System;
using Aspose.Cells;

// The sample creates a workbook, fills columns A and B with numeric values and criteria, inserts a MINIFS formula in C1 to find the minimum of A1:A5 where B1:B5 equals "Yes", sets the workbook to Excel 2016 compatibility mode, calculates all formulas, retrieves the result, compares it to the expected value (8), and prints whether the MINIFS function returned the correct result, handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            double[] values = { 10, 5, 8, 3, 12 };
            string[] criteria = { "Yes", "No", "Yes", "No", "Yes" };
            for (int i = 0; i < values.Length; i++)
            {
                cells[i, 0].PutValue(values[i]);   // Column A
                cells[i, 1].PutValue(criteria[i]); // Column B
            }

            // Insert MINIFS formula: find minimum of A1:A5 where B1:B5 = "Yes"
            cells["C1"].Formula = "MINIFS(A1:A5,B1:B5,\"Yes\")";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the calculated result
            double result = cells["C1"].DoubleValue;

            // Expected result based on the data (minimum of 10, 8, 12)
            double expected = 8;

            // Validate the result
            if (Math.Abs(result - expected) < 1e-9)
            {
                Console.WriteLine("MINIFS function returned correct result: " + result);
            }
            else
            {
                Console.WriteLine($"MINIFS function returned {result}, expected {expected}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
