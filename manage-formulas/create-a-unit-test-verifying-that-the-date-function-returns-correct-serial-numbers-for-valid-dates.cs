// Title: C# unit test with Aspose.Cells to confirm Excel DATE function returns the correct serial number
// Common Searches: how to write a unit test for Aspose.Cells DATE formula in C# | verify Excel DATE function serial value using Aspose.Cells CalculateFormula | C# Aspose.Cells test DATE(year,month,day) matches .NET OADate | unit testing Excel date serial numbers with Aspose.Cells library | Aspose.Cells validate DATE function output against expected serial number
// Tags: Aspose.Cells DATE formula verification | Excel OADate serial number comparison | Workbook.CalculateFormula usage example | Aspose.Cells formula evaluation in C# | date serial number assertion

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a workbook, inserts a DATE(year,month,day) formula into cell A1, runs Workbook.CalculateFormula, and asserts that the resulting double matches the .NET OADate value within a small tolerance, throwing an exception on mismatch.
    public class DateFunctionTests
    {
        public static void Main()
        {
            try
            {
                RunDateFunctionTest();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        private static void RunDateFunctionTest()
        {
            // Arrange: create a new workbook and access the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            // Define a test date
            int year = 2023;
            int month = 5;
            int day = 10;

            // Expected Excel serial number for the date (using OLE Automation date conversion)
            double expectedSerial = new DateTime(year, month, day).ToOADate();

            // Place the DATE function formula in cell A1
            cells["A1"].Formula = $"=DATE({year},{month},{day})";

            // Act: calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the calculated value from cell A1
            double actualSerial = cells["A1"].DoubleValue;

            // Assert: the actual serial number should match the expected value within tolerance
            const double tolerance = 0.000001;
            if (Math.Abs(expectedSerial - actualSerial) > tolerance)
            {
                throw new InvalidOperationException(
                    $"DATE({year},{month},{day}) returned {actualSerial} but expected {expectedSerial}.");
            }
        }
    }
}
