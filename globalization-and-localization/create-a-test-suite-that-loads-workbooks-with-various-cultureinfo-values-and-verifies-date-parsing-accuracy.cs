// Title: Create a C# Aspose.Cells test suite that verifies date parsing accuracy for multiple CultureInfo locales
// AI Prompts: Generate an in‑memory XLSX workbook containing a specific date string, load it with Aspose.Cells using LoadOptions.CultureInfo for a given locale, and assert that the cell is recognized as a DateTime matching the expected value. | Implement a parameterized NUnit test that iterates over a collection of culture/date‑string pairs, creates the workbook, loads it with the appropriate CultureInfo, and logs pass/fail results for each case. | Extend the suite to handle custom date formats by configuring LoadOptions.NumberFormat and provide detailed diagnostic output when a parsing mismatch occurs.
// Common Searches: how to test date parsing with different cultures using Aspose.Cells in C# | Aspose.Cells LoadOptions CultureInfo example for Excel date conversion | C# unit test for Excel workbook date localization with Aspose.Cells | verify that Aspose.Cells interprets date strings correctly for en-GB and de-DE | load Excel file with specific locale to parse dates in Aspose.Cells .NET
// Tags: Aspose.Cells culture-aware workbook loading | Excel date string parsing validation with Aspose.Cells | C# in‑memory XLSX generation for localization testing | unit test for multi‑locale date conversion in .NET | LoadOptions.CultureInfo usage for Excel date handling

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCultureTests
{
    // The example creates in‑memory XLSX workbooks with a single date string cell, loads each workbook using Aspose.Cells LoadOptions configured with a specific CultureInfo, checks that the cell type is DateTime, compares the parsed date to the expected value, and reports pass/fail results for several locales.
    public class DateParsingTests
    {
        // Creates a workbook with a single cell (A1) containing the supplied date string.
        private MemoryStream CreateWorkbookWithDateString(string dateString)
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue(dateString); // store as string
            var stream = new MemoryStream();
            wb.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0;
            return stream;
        }

        // Loads the workbook using the specified culture and verifies the parsed DateTime.
        private void VerifyDateParsing(string cultureName, string dateString, DateTime expectedDate)
        {
            try
            {
                using (MemoryStream stream = CreateWorkbookWithDateString(dateString))
                {
                    // Configure load options with the target culture.
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        CultureInfo = new CultureInfo(cultureName)
                    };

                    // Load workbook using the culture-aware options.
                    var wb = new Workbook(stream, loadOptions);
                    var ws = wb.Worksheets[0];
                    var cell = ws.Cells["A1"];

                    // Ensure the cell was interpreted as a DateTime.
                    if (cell.Type != CellValueType.IsDateTime)
                    {
                        Console.WriteLine($"[FAIL] Culture {cultureName}: Cell type is not DateTime (actual: {cell.Type}).");
                        return;
                    }

                    // Compare the parsed date with the expected value (ignoring time component).
                    DateTime actualDate = cell.DateTimeValue.Date;
                    if (actualDate != expectedDate.Date)
                    {
                        Console.WriteLine($"[FAIL] Culture {cultureName}: Expected {expectedDate:yyyy-MM-dd}, got {actualDate:yyyy-MM-dd}.");
                    }
                    else
                    {
                        Console.WriteLine($"[PASS] Culture {cultureName}: Parsed correctly as {actualDate:yyyy-MM-dd}.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Culture {cultureName}: {ex.Message}");
            }
        }

        // Executes all test cases.
        public void RunAllTests()
        {
            var testCases = new[]
            {
                new { Culture = "en-US", DateStr = "12/31/2023", Year = 2023, Month = 12, Day = 31 },
                new { Culture = "en-GB", DateStr = "31/12/2023", Year = 2023, Month = 12, Day = 31 },
                new { Culture = "fr-FR", DateStr = "31/12/2023", Year = 2023, Month = 12, Day = 31 },
                new { Culture = "de-DE", DateStr = "31.12.2023", Year = 2023, Month = 12, Day = 31 },
                new { Culture = "ja-JP", DateStr = "2023/12/31", Year = 2023, Month = 12, Day = 31 }
            };

            foreach (var tc in testCases)
            {
                DateTime expected = new DateTime(tc.Year, tc.Month, tc.Day);
                VerifyDateParsing(tc.Culture, tc.DateStr, expected);
            }
        }

        // Entry point.
        public static void Main(string[] args)
        {
            try
            {
                var tester = new DateParsingTests();
                tester.RunAllTests();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
