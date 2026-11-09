// Title: C# test suite for verifying Aspose.Cells CSV date parsing with various CultureInfo settings
// Common Searches: how to unit test Aspose.Cells CSV date parsing for en-US and en-GB cultures in C# | load a CSV into Aspose.Cells workbook with a specific CultureInfo and check the DateTime value | C# example of testing culture‑specific date formats using Aspose.Cells LoadOptions | verify that Aspose.Cells respects .NET CultureInfo when reading dates from CSV files | automated test for multicultural date conversion in Aspose.Cells using in‑memory streams
// Tags: Aspose.Cells CSV culture-specific date parsing | C# unit test for workbook loading with LoadOptions | Validate DateTime conversion in Aspose.Cells worksheets | CultureInfo handling in Aspose.Cells CSV import | In‑memory CSV stream testing with Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCultureTests
{
    // The example defines a set of culture‑specific test cases, creates an in‑memory CSV containing a date string, loads it into an Aspose.Cells Workbook via LoadOptions, assigns the desired CultureInfo, extracts the cell value, parses it to DateTime, and asserts that the parsed date matches the expected value, outputting pass/fail results for each culture.
    class Program
    {
        // Entry point that runs the date‑parsing examples.
        static void Main()
        {
            // Define test cases: culture, CSV date string, expected year, month, day.
            var testCases = new (string culture, string csvDate, int year, int month, int day)[]
            {
                ("en-US", "01/02/2020", 2020, 1, 2), // MM/dd/yyyy => Jan 2, 2020
                ("en-GB", "01/02/2020", 2020, 2, 1), // dd/MM/yyyy => 1 Feb 2020
                ("fr-FR", "01/02/2020", 2020, 2, 1), // dd/MM/yyyy => 1 Feb 2020
                ("de-DE", "01/02/2020", 2020, 2, 1)  // dd/MM/yyyy accepted => 1 Feb 2020
            };

            foreach (var (culture, csvDate, year, month, day) in testCases)
            {
                try
                {
                    DateTime expected = new DateTime(year, month, day);
                    TestDateParsingByCulture(culture, csvDate, expected);
                    Console.WriteLine($"[PASS] Culture '{culture}' correctly parsed '{csvDate}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[FAIL] Culture '{culture}' failed to parse '{csvDate}': {ex.Message}");
                }
            }
        }

        // Creates a CSV stream containing a single cell with the supplied date string.
        private static MemoryStream CreateCsvWithDateString(string dateString)
        {
            var stream = new MemoryStream();
            using (var writer = new StreamWriter(stream, leaveOpen: true))
            {
                writer.Write(dateString);
                writer.Flush();
                stream.Position = 0;
            }
            return stream;
        }

        // Verifies that the date parsed from the workbook matches the expected DateTime for a given culture.
        private static void VerifyDateParsing(string cultureName, string csvDateString, DateTime expectedDate)
        {
            // Load CSV data into a workbook.
            using (MemoryStream csvStream = CreateCsvWithDateString(csvDateString))
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
                Workbook workbook;
                try
                {
                    workbook = new Workbook(csvStream, loadOptions);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Failed to load CSV into workbook.", ex);
                }

                // Apply the culture before reading the cell value.
                workbook.Settings.CultureInfo = new CultureInfo(cultureName);

                Cell cell = workbook.Worksheets[0].Cells[0, 0];
                object cellValue = cell.Value;

                DateTime parsedDate;
                if (cellValue is DateTime dt)
                {
                    parsedDate = dt;
                }
                else
                {
                    try
                    {
                        parsedDate = DateTime.Parse(cellValue?.ToString() ?? string.Empty, workbook.Settings.CultureInfo);
                    }
                    catch (Exception ex)
                    {
                        throw new FormatException($"Unable to parse cell value '{cellValue}' using culture '{cultureName}'.", ex);
                    }
                }

                if (parsedDate != expectedDate)
                {
                    throw new InvalidOperationException(
                        $"Culture '{cultureName}' parsed '{csvDateString}' as {parsedDate:d}, expected {expectedDate:d}.");
                }
            }
        }

        // Wrapper that prepares expected DateTime and calls verification.
        private static void TestDateParsingByCulture(string cultureName, string csvDateString, DateTime expected)
        {
            VerifyDateParsing(cultureName, csvDateString, expected);
        }
    }
}
