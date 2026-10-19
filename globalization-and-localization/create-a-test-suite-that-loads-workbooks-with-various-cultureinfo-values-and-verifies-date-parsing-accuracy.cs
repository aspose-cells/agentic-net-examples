// Title: C# Aspose.Cells test suite that validates Excel date parsing for multiple CultureInfo locales
// Common Searches: how to unit test Aspose.Cells date parsing with different CultureInfo settings in C# | load Excel workbook with specific CultureInfo using Aspose.Cells LoadOptions | verify that Aspose.Cells correctly parses localized date strings in XLSX files | C# example for testing date string conversion across en-US, fr-FR, de-DE cultures with Aspose.Cells | Aspose.Cells culture-aware date parsing validation in automated tests
// Tags: Aspose.Cells LoadOptions CultureInfo date parsing | C# in‑memory workbook date string test | Excel localized date format verification Aspose.Cells | unit testing Excel date handling with multiple locales | culture‑specific date parsing validation Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCultureTests
{
    // The example creates in‑memory XLSX workbooks containing a date string formatted for each specified CultureInfo, loads each workbook with LoadOptions.CultureInfo set to the same locale, and asserts that Aspose.Cells interprets the cell as a DateTime with the expected date component, enabling automated verification of culture‑aware date parsing.
    class Program
    {
        // Helper: creates a workbook with a date string in cell A1 and returns the stream.
        private static MemoryStream CreateWorkbookWithDateString(string dateString)
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue(dateString); // store as string
            var stream = new MemoryStream();
            wb.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0;
            return stream;
        }

        static void Main()
        {
            try
            {
                // Date to be written and later verified.
                DateTime originalDate = new DateTime(2023, 5, 17, 14, 30, 0);

                // Cultures to test.
                string[] cultureNames = { "en-US", "fr-FR", "de-DE", "ja-JP", "es-ES" };

                foreach (string cultureName in cultureNames)
                {
                    try
                    {
                        CultureInfo culture = new CultureInfo(cultureName);

                        // Format the date according to the current culture (short date pattern).
                        string dateString = originalDate.ToString(culture);

                        // Create workbook containing the formatted date string.
                        using (MemoryStream stream = CreateWorkbookWithDateString(dateString))
                        {
                            // LoadOptions with the same culture for parsing.
                            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                            {
                                CultureInfo = culture
                            };

                            // Load the workbook.
                            Workbook loadedWb = new Workbook(stream, loadOptions);
                            Cell loadedCell = loadedWb.Worksheets[0].Cells["A1"];

                            // Verify that Aspose.Cells parsed the string into a DateTime.
                            if (loadedCell.Type != CellValueType.IsDateTime)
                            {
                                Console.WriteLine($"[FAIL] Cell type is not DateTime for culture {cultureName}.");
                                continue;
                            }

                            DateTime parsedDate = loadedCell.DateTimeValue;

                            // Compare only the date part (time may be lost depending on format).
                            if (originalDate.Date != parsedDate.Date)
                            {
                                Console.WriteLine($"[FAIL] Parsed date does not match for culture {cultureName}.");
                            }
                            else
                            {
                                Console.WriteLine($"[PASS] Culture {cultureName}: parsed date matches.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[ERROR] Exception for culture {cultureName}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
