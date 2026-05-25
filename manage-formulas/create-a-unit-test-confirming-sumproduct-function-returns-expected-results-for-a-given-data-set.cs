using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    public class SumProductFunctionTests
    {
        public void SumProduct_ReturnsExpectedResult()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cells = sheet.Cells;

                // Populate first array: A1:A3 = {1, 2, 3}
                cells["A1"].PutValue(1);
                cells["A2"].PutValue(2);
                cells["A3"].PutValue(3);

                // Populate second array: B1:B3 = {4, 5, 6}
                cells["B1"].PutValue(4);
                cells["B2"].PutValue(5);
                cells["B3"].PutValue(6);

                // Set SUMPRODUCT formula in C1
                // Expected result: 1*4 + 2*5 + 3*6 = 32
                cells["C1"].Formula = "=SUMPRODUCT(A1:A3,B1:B3)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Retrieve the calculated value
                object result = cells["C1"].Value;

                // Verify the result is 32
                if (Math.Abs(Convert.ToDouble(result) - 32.0) > 0.0001)
                {
                    throw new InvalidOperationException($"Unexpected result: {result}");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "SumProductFunctionTest.xlsx";
                try
                {
                    workbook.Save(outputPath);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Warning: Could not save workbook to '{outputPath}'. {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Propagate exception to caller for handling/logging
                throw new ApplicationException("SumProduct test failed.", ex);
            }
        }
    }

    internal class Program
    {
        static void Main()
        {
            try
            {
                var test = new SumProductFunctionTests();
                test.SumProduct_ReturnsExpectedResult();
                Console.WriteLine("SumProduct test passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
            }
        }
    }
}