using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    public class DateFunctionTests
    {
        // Compute expected Excel serial number for a DateTime using the workbook's date system
        private static double GetExpectedSerial(DateTime date, bool date1904)
        {
            return CellsHelper.GetDoubleFromDateTime(date, date1904);
        }

        public void Run()
        {
            // Test data: (input year, month, day, expected normalized year, month, day)
            var testCases = new (int year, int month, int day, int expYear, int expMonth, int expDay)[]
            {
                (2021, 1, 1, 2021, 1, 1),          // Simple date
                (1900, 1, 1, 1900, 1, 1),          // Excel epoch (1900 system)
                (1904, 1, 1, 1904, 1, 1),          // Excel epoch (1904 system)
                (2020, 2, 29, 2020, 2, 29),        // Leap day
                (2021, 13, 1, 2022, 1, 1),         // Month overflow (13 -> Jan of next year)
                (2021, 0, 15, 2020, 12, 15),       // Month underflow (0 -> Dec of previous year)
                (2021, 1, 32, 2021, 2, 1)          // Day overflow (32 Jan -> 1 Feb)
            };

            foreach (var tc in testCases)
            {
                try
                {
                    // Arrange: create a workbook and set the DATE formula
                    var workbook = new Workbook();
                    var sheet = workbook.Worksheets[0];
                    var cell = sheet.Cells["A1"];
                    cell.Formula = $"=DATE({tc.year},{tc.month},{tc.day})";

                    // Act: calculate formulas
                    workbook.CalculateFormula();

                    // Retrieve the calculated serial number
                    double actualSerial = cell.DoubleValue;

                    // Expected DateTime after Excel normalizes the inputs
                    var expectedDate = new DateTime(tc.expYear, tc.expMonth, tc.expDay);

                    // Determine which date system the workbook uses (default is 1900)
                    bool date1904 = workbook.Settings.Date1904;

                    // Compute expected serial using CellsHelper
                    double expectedSerial = GetExpectedSerial(expectedDate, date1904);

                    // Assert: compare with a tiny tolerance
                    if (Math.Abs(expectedSerial - actualSerial) > 1e-9)
                    {
                        Console.WriteLine($"FAIL: DATE({tc.year},{tc.month},{tc.day}) => {actualSerial}, expected {expectedSerial}");
                    }
                    else
                    {
                        Console.WriteLine($"PASS: DATE({tc.year},{tc.month},{tc.day}) => {actualSerial}");
                    }
                }
                catch (Exception ex)
                {
                    // Runtime safety: report any unexpected errors per test case
                    Console.WriteLine($"ERROR: DATE({tc.year},{tc.month},{tc.day}) threw an exception: {ex.Message}");
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                var tests = new DateFunctionTests();
                tests.Run();
            }
            catch (Exception ex)
            {
                // Catch any unhandled exceptions
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}