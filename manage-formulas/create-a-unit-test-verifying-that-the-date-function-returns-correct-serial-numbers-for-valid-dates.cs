// Title: C# unit test for Aspose.Cells DATE formula verifying Excel serial date values
// Common Searches: how to write a C# unit test for Aspose.Cells DATE formula | verify Excel serial numbers returned by Aspose.Cells DATE function in .NET | C# example testing DATE function leap year handling with Aspose.Cells | Aspose.Cells calculate formula DATE returns correct OADate value | unit testing Excel date conversion using Aspose.Cells in C#
// Tags: Aspose.Cells DATE formula unit testing | Excel serial date verification C# | OADate conversion test Aspose.Cells | leap year handling Aspose.Cells DATE | C# workbook formula evaluation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a workbook, assigns the DATE formula to a cell for several dates (including a leap year), recalculates the workbook, extracts the resulting serial number via DoubleValue, and asserts that it matches the .NET OADate value, outputting PASS or FAIL for each test case.
    public class DateFunctionTests
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cells = sheet.Cells;

                // Test data: year, month, day.
                var testDates = new[]
                {
                    new { Year = 1900, Month = 1, Day = 1 },
                    new { Year = 1900, Month = 2, Day = 28 },
                    new { Year = 1900, Month = 3, Day = 1 },
                    new { Year = 1999, Month = 12, Day = 31 },
                    new { Year = 2000, Month = 2, Day = 29 }, // leap year
                    new { Year = 2023, Month = 1, Day = 1 }
                };

                foreach (var td in testDates)
                {
                    // Set the DATE formula in cell A1.
                    cells["A1"].Formula = $"=DATE({td.Year},{td.Month},{td.Day})";

                    // Recalculate the workbook to evaluate the formula.
                    workbook.CalculateFormula();

                    // Retrieve the calculated serial number.
                    double actualSerial = cells["A1"].DoubleValue;

                    // Expected serial number using .NET OADate (matches Excel's serial system).
                    double expectedSerial = new DateTime(td.Year, td.Month, td.Day).ToOADate();

                    // Verify that the serial numbers match.
                    if (Math.Abs(expectedSerial - actualSerial) > 1e-10)
                    {
                        Console.WriteLine($"FAIL: DATE({td.Year},{td.Month},{td.Day}) returned {actualSerial}, expected {expectedSerial}.");
                    }
                    else
                    {
                        Console.WriteLine($"PASS: DATE({td.Year},{td.Month},{td.Day}) = {actualSerial}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
