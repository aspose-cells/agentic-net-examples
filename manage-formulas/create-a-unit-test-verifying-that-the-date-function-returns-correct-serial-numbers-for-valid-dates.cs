// Title: Unit Test for Excel DATE Function Serial Numbers with Aspose.Cells (.NET)
// Description: A C# example that defines a set of DateTime values and their expected Excel serial numbers under the default 1900 date system, uses CellsHelper.GetDoubleFromDateTime to convert each date, and asserts the result matches the expected value within a 1e‑6 tolerance. The test covers the Excel leap‑year bug and fractional time values, throwing detailed errors on mismatches.
// Keywords: Aspose.Cells | C# unit test | Excel DATE serial number | CellsHelper.GetDoubleFromDateTime | 1900 date system | Excel leap year bug | .NET date conversion | date to serial conversion | unit testing Excel dates | Aspose.Cells example
// Common Searches: Aspose.Cells unit test DATE function | verify Excel date serial numbers C# | GetDoubleFromDateTime usage example | test 1900 date system Aspose.Cells | Excel leap year bug unit test | C# test for Excel date conversion
// Developer Intent: Confirm that CellsHelper.GetDoubleFromDateTime returns the correct Excel serial numbers for given DateTime inputs, including handling of the 1900 leap‑year anomaly and time fractions.
// Use Cases: Validate that spreadsheet calculations relying on the DATE function produce accurate results. | Detect regressions in Aspose.Cells date handling after library upgrades. | Demonstrate proper testing of the 1900 date system and its known leap‑year bug. | Ensure time components are correctly represented as fractional parts of the serial number. | Provide a reusable pattern for unit testing date conversions in .NET projects.
// AI Prompts: Generate an MSTest method that asserts CellsHelper.GetDoubleFromDateTime returns expected serial numbers for a list of dates, handling the 1900 leap‑year bug. | Create a NUnit test class that iterates over DateTime/expected serial pairs with a tolerance of 0.000001 and reports detailed error messages. | Write an xUnit Theory test for verifying Excel date serial conversion using Aspose.Cells, covering whole dates and date‑time values. | Produce documentation code snippets showing how to test Excel DATE function serial numbers with Aspose.Cells in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // A C# example that defines a set of DateTime values and their expected Excel serial numbers under the default 1900 date system, uses CellsHelper.GetDoubleFromDateTime to convert each date, and asserts the result matches the expected value within a 1e‑6 tolerance. The test covers the Excel leap‑year bug and fractional time values, throwing detailed errors on mismatches.
    public class Program
    {
        public static void Main()
        {
            try
            {
                VerifyDateFunctionSerialNumbers();
                Console.WriteLine("All date serial number checks passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void VerifyDateFunctionSerialNumbers()
        {
            // Use the default 1900 date system (false)
            bool date1904 = false;

            // Define test dates and their expected Excel serial numbers
            var testData = new (DateTime date, double expected)[]
            {
                (new DateTime(1900, 1, 1), 1),          // Excel day 1
                (new DateTime(1900, 1, 2), 2),          // Excel day 2
                (new DateTime(1900, 2, 28), 59),        // Day before the fictitious leap day
                (new DateTime(1900, 3, 1), 61),         // Skips 60 due to Excel's leap year bug
                (new DateTime(2021, 1, 1), 44197),      // Known serial for 2021‑01‑01
                (new DateTime(2021, 1, 1, 12, 0, 0), 44197.5) // Same date with 12:00 PM
            };

            foreach (var (date, expected) in testData)
            {
                try
                {
                    // Convert DateTime to Excel serial number
                    double serial = CellsHelper.GetDoubleFromDateTime(date, date1904);

                    // Verify the conversion is accurate (allowing a tiny tolerance)
                    if (Math.Abs(expected - serial) > 0.000001)
                    {
                        throw new InvalidOperationException(
                            $"Serial number for {date:yyyy-MM-dd HH:mm:ss} is incorrect. Expected {expected}, got {serial}.");
                    }
                }
                catch (Exception ex)
                {
                    // Rethrow with context to be caught by outer handler
                    throw new Exception($"Error processing date {date:yyyy-MM-dd HH:mm:ss}: {ex.Message}", ex);
                }
            }
        }
    }
}
