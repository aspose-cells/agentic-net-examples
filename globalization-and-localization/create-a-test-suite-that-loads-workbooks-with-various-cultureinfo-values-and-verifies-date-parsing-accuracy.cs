// Title: C# Test Suite for Verifying Date Parsing with Multiple CultureInfo Settings in Aspose.Cells
// Description: A self‑contained C# program that creates a workbook with an ISO date string and an Excel serial date, saves it, then reloads the file using LoadOptions.CultureInfo for en‑US, de‑DE and fr‑FR. The suite validates that Cells["A1"].DateTimeValue and Cells["A2"].DateTimeValue both equal 31 Dec 2020 and confirms that Workbook.Settings.CultureInfo matches the culture supplied at load time.
// Keywords: Aspose.Cells | C# | CultureInfo | date parsing | LoadOptions.CultureInfo | Workbook.Settings.CultureInfo | Excel serial date | unit test | localization | globalization
// Common Searches: Aspose.Cells test date parsing CultureInfo | load Excel workbook with specific CultureInfo .NET | verify DateTimeValue across locales Aspose | unit test Excel serial date in different cultures | Workbook.Settings.CultureInfo example
// Developer Intent: Ensure Aspose.Cells correctly interprets both string and numeric dates when a workbook is loaded with various CultureInfo configurations.
// Use Cases: Automated regression test that guarantees consistent date values for en‑US, de‑DE, and fr‑FR locales. | CI pipeline check that the culture passed to LoadOptions is reflected in Workbook.Settings.CultureInfo. | Localization QA script to confirm that date formatting does not alter underlying DateTime data in exported reports.
// AI Prompts: Generate NUnit test methods that iterate over a list of CultureInfo objects, load the sample workbook with LoadOptions.CultureInfo, and assert that Cells["A1"].DateTimeValue and Cells["A2"].DateTimeValue equal 2020‑12‑31. | Write a PowerShell script to compile and run the provided C# test suite on a build server, capturing PASS/FAIL results for each culture. | Create a reusable helper class that abstracts workbook creation, culture‑specific loading, and date verification for Aspose.Cells unit tests.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCultureInfoTests
{
    // Test suite for verifying date parsing with different CultureInfo settings.
    // A self‑contained C# program that creates a workbook with an ISO date string and an Excel serial date, saves it, then reloads the file using LoadOptions.CultureInfo for en‑US, de‑DE and fr‑FR. The suite validates that Cells["A1"].DateTimeValue and Cells["A2"].DateTimeValue both equal 31 Dec 2020 and confirms that Workbook.Settings.CultureInfo matches the culture supplied at load time.
    public class TestCultureInfoDateParsing
    {
        // Path for the temporary workbook used in tests.
        private const string SampleFilePath = "SampleDateWorkbook.xlsx";

        // Entry point.
        public static void Main()
        {
            try
            {
                // Prepare a workbook containing a date string and a numeric date.
                CreateSampleWorkbook();

                // Define cultures to test.
                var cultures = new List<CultureInfo>
                {
                    new CultureInfo("en-US"), // MM/dd/yyyy
                    new CultureInfo("de-DE"), // dd.MM.yyyy
                    new CultureInfo("fr-FR")  // dd/MM/yyyy
                };

                // Expected DateTime value (31 December 2020).
                DateTime expectedDate = new DateTime(2020, 12, 31);

                // Run the verification for each culture.
                foreach (var culture in cultures)
                {
                    VerifyDateParsing(culture, expectedDate);
                }

                Console.WriteLine("All tests completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Creates a workbook with a date string in cell A1 and a numeric date in cell A2.
        private static void CreateSampleWorkbook()
        {
            try
            {
                // Create a new workbook.
                Workbook wb = new Workbook();

                // Worksheet reference.
                Worksheet sheet = wb.Worksheets[0];

                // Cell A1: date string in ISO format (unambiguous for parsing).
                sheet.Cells["A1"].PutValue("2020-12-31");

                // Cell A2: Excel serial number for the same date.
                // Excel's base date is 1899-12-30 for the 1900 date system.
                double excelSerial = wb.Settings.Date1904 ? 43831 : 44197; // Adjusted for 1900 system.
                sheet.Cells["A2"].PutValue(excelSerial);

                // Apply custom number format to A2.
                Style style = sheet.Cells["A2"].GetStyle();
                style.Custom = "mm/dd/yyyy";
                sheet.Cells["A2"].SetStyle(style);

                // Save the workbook to disk.
                wb.Save(SampleFilePath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating sample workbook: {ex.Message}");
                throw;
            }
        }

        // Loads the workbook with the specified culture and verifies date values.
        private static void VerifyDateParsing(CultureInfo culture, DateTime expected)
        {
            try
            {
                if (!File.Exists(SampleFilePath))
                {
                    Console.WriteLine($"File not found: {SampleFilePath}");
                    return;
                }

                // Configure load options with the target CultureInfo.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    CultureInfo = culture // Use the rule LoadOptions.CultureInfo
                };

                // Load the workbook using the configured options.
                Workbook wb = new Workbook(SampleFilePath, loadOptions);

                // Access the first worksheet.
                Worksheet sheet = wb.Worksheets[0];

                // Verify cell A1 (string date) is parsed to the expected DateTime.
                DateTime parsedFromString = sheet.Cells["A1"].DateTimeValue;
                bool stringParseOk = parsedFromString.Date == expected.Date;

                // Verify cell A2 (numeric date) is interpreted correctly.
                DateTime parsedFromNumber = sheet.Cells["A2"].DateTimeValue;
                bool numberParseOk = parsedFromNumber.Date == expected.Date;

                // Output results.
                Console.WriteLine($"Culture: {culture.Name}");
                Console.WriteLine($"  A1 parsed date: {parsedFromString:d} - {(stringParseOk ? "PASS" : "FAIL")}");
                Console.WriteLine($"  A2 parsed date: {parsedFromNumber:d} - {(numberParseOk ? "PASS" : "FAIL")}");

                // Additional check: WorkbookSettings.CultureInfo reflects the loaded culture.
                CultureInfo settingsCulture = wb.Settings.CultureInfo;
                bool settingsMatch = settingsCulture != null &&
                                     settingsCulture.Name.Equals(culture.Name, StringComparison.OrdinalIgnoreCase);
                Console.WriteLine($"  WorkbookSettings.CultureInfo matches: {(settingsMatch ? "YES" : "NO")}");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error verifying culture {culture.Name}: {ex.Message}");
            }
        }
    }
}
