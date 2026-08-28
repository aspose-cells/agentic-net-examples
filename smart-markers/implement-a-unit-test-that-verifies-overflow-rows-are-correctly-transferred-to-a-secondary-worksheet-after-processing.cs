// Title: C# unit test to verify overflow rows are moved to a secondary worksheet using TxtLoadOptions in Aspose.Cells
// AI Prompts: Generate a C# unit test that loads CSV data with TxtLoadOptions (MaxRowCount = 5, ExtendToNextSheet = true, HeaderRowsCount = 1) and asserts that a second worksheet is created containing the overflow rows with the expected cell values. | Write assertions to confirm that the header row stays on the first sheet, the row counts on both worksheets are correct, and the first and last overflow rows match the original CSV data.
// Common Searches: c# Aspose.Cells test overflow rows to another sheet using TxtLoadOptions | how to assert second worksheet creation after CSV exceeds MaxRowCount in Aspose.Cells | unit testing CSV import with ExtendToNextSheet option in Aspose.Cells .NET | verify header row handling with TxtLoadOptions.HeaderRowsCount in Aspose.Cells
// Tags: TxtLoadOptions overflow rows to next worksheet | Aspose.Cells CSV import MaxRowCount | ExtendToNextSheet worksheet creation Aspose.Cells | HeaderRowsCount validation Aspose.Cells | C# unit test Aspose.Cells worksheet overflow

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Demonstrates a C# unit test that loads CSV data using TxtLoadOptions (MaxRowCount=5, ExtendToNextSheet=true, HeaderRowsCount=1) and validates that overflow rows are placed on a second worksheet, checking worksheet count, row indices, and specific cell values.
    public class TxtLoadOptionsOverflowDemo
    {
        public static void Main()
        {
            try
            {
                // Prepare CSV data with 10 rows (including header)
                var sb = new StringBuilder();
                sb.AppendLine("Header1,Header2");
                for (int i = 1; i <= 9; i++)
                {
                    sb.AppendLine($"Data{i}A,Data{i}B");
                }
                string csvContent = sb.ToString();

                // Configure TxtLoadOptions:
                // - MaxRowCount = 5 (first sheet holds 5 rows)
                // - ExtendToNextSheet = true (overflow rows go to a new sheet)
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    MaxRowCount = 5,
                    ExtendToNextSheet = true,
                    HeaderRowsCount = 1 // first row is header
                };

                // Load the CSV into a workbook using the options
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(csvContent)))
                {
                    Workbook workbook = new Workbook(ms, loadOptions);

                    // Verify that a second worksheet was created
                    if (workbook.Worksheets.Count != 2)
                        throw new InvalidOperationException("A second worksheet should be created for overflow rows.");

                    // First worksheet should contain rows 0..4 (header + 4 data rows)
                    Worksheet sheet1 = workbook.Worksheets[0];
                    Cells cells1 = sheet1.Cells;
                    if (cells1.MaxDataRow != 4)
                        throw new InvalidOperationException("First sheet should have 5 rows (0‑based index 0‑4).");
                    if (cells1[4, 0].StringValue != "Data4A" || cells1[4, 1].StringValue != "Data4B")
                        throw new InvalidOperationException("Content of the last row in the first sheet is incorrect.");

                    // Second worksheet should contain the remaining rows (rows 5..9 of original data)
                    Worksheet sheet2 = workbook.Worksheets[1];
                    Cells cells2 = sheet2.Cells;
                    if (cells2.MaxDataRow != 4)
                        throw new InvalidOperationException("Second sheet should have 5 rows (0‑based index 0‑4).");
                    if (cells2[0, 0].StringValue != "Data5A" || cells2[0, 1].StringValue != "Data5B")
                        throw new InvalidOperationException("First row of the second sheet is incorrect.");
                    if (cells2[4, 0].StringValue != "Data9A" || cells2[4, 1].StringValue != "Data9B")
                        throw new InvalidOperationException("Last row of the second sheet is incorrect.");

                    Console.WriteLine("All checks passed successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
