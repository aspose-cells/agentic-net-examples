// Title: Unit Test for Aspose.Cells DATE Function – Verify Excel Serial Number in C#
// Description: Demonstrates how to create a C# unit test that inserts the DATE(2021,1,1) formula into a workbook, evaluates it with Aspose.Cells, retrieves the resulting serial number, computes the expected value using CellsHelper.GetDoubleFromDateTime (respecting the workbook's date system), and asserts equality within a 1e-6 tolerance.
// Keywords: Aspose.Cells | DATE function | unit test | C# | .NET | Excel serial date | CellsHelper | CalculateFormula | date1904 | formula evaluation
// Common Searches: Aspose.Cells unit test DATE formula | How to test Excel date serial number with Aspose.Cells | C# verify DATE function returns correct serial | Aspose.Cells GetDoubleFromDateTime example | Test DATE function 1904 date system Aspose
// Developer Intent: Confirm that Aspose.Cells DATE formula returns the correct Excel serial number.
// Use Cases: Automated regression testing of date‑related formulas. | Validate workbook date‑system conversions (1900 vs 1904). | Ensure consistency when migrating spreadsheets between platforms. | Generate test data for documentation of date handling. | Check custom date logic in add‑ins or extensions.
// AI Prompts: Create an MSTest method that asserts DATE(2021,1,1) serial matches CellsHelper.GetDoubleFromDateTime with a 1e-6 tolerance. | Write an NUnit test covering both 1900 and 1904 date systems for the DATE function in Aspose.Cells. | Provide xUnit code to compare DATE(2022,12,31) result with the expected serial using CellsHelper. | Generate a parameterized test that verifies DATE(year, month, day) across a range of dates in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Demonstrates how to create a C# unit test that inserts the DATE(2021,1,1) formula into a workbook, evaluates it with Aspose.Cells, retrieves the resulting serial number, computes the expected value using CellsHelper.GetDoubleFromDateTime (respecting the workbook's date system), and asserts equality within a 1e-6 tolerance.
    public class DateFunctionTests
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Insert the DATE formula for a known date (2021-01-01)
                worksheet.Cells["A1"].Formula = "DATE(2021,1,1)";

                // Evaluate the formula
                workbook.CalculateFormula();

                // Retrieve the serial number produced by the DATE function
                double actualSerial = worksheet.Cells["A1"].DoubleValue;

                // Compute the expected serial number using CellsHelper (same date system)
                DateTime expectedDate = new DateTime(2021, 1, 1);
                bool date1904 = workbook.Settings.Date1904; // default is false (1900 system)
                double expectedSerial = CellsHelper.GetDoubleFromDateTime(expectedDate, date1904);

                // Verify that the serial numbers match within a small tolerance
                const double tolerance = 1e-6;
                if (Math.Abs(expectedSerial - actualSerial) <= tolerance)
                {
                    Console.WriteLine($"Test passed. Serial number: {actualSerial}");
                }
                else
                {
                    Console.WriteLine($"Test failed. Expected {expectedSerial}, got {actualSerial}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
        }
    }
}
