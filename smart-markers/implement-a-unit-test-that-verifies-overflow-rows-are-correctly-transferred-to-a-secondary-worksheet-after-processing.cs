// Title: C# unit test for TxtLoadOptions ExtendToNextSheet overflow handling in Aspose.Cells
// Description: Creates a CSV with a header and nine rows, loads it with TxtLoadOptions (MaxRowCount = 5, ExtendToNextSheet = true, HeaderRowsCount = 1) and verifies that two worksheets are generated, the first contains the header plus rows 1‑4, the second contains rows 5‑9, and that cells beyond the expected range remain empty.
// Keywords: Aspose.Cells | TxtLoadOptions | ExtendToNextSheet | MaxRowCount | C# unit test | CSV overflow to new sheet | Workbook worksheet count | cell validation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells unit test ExtendToNextSheet | How to verify overflow rows in Aspose.Cells | TxtLoadOptions MaxRowCount example C# | CSV load creates second worksheet Aspose | Validate empty cells after overflow in Aspose.Cells
// Developer Intent: Write an automated test that confirms rows exceeding the MaxRowCount limit are automatically moved to a newly created worksheet when ExtendToNextSheet is enabled.
// Use Cases: Ensure CSV imports respect MaxRowCount and split excess rows into a separate sheet. | Validate that the header row stays on the first worksheet while overflow data starts on the second. | Check that no unintended data appears beyond the defined row range in either worksheet.
// AI Prompts: Generate an MSTest/NUnit/XUnit test that loads a CSV stream with TxtLoadOptions (MaxRowCount=5, ExtendToNextSheet=true) and asserts worksheet count, header placement, and row values. | Create mock CSV data and a C# unit test to verify that overflow rows are transferred to a new worksheet using Aspose.Cells. | Explain how to programmatically confirm that cells beyond the expected rows (e.g., A6) are empty in both primary and overflow worksheets after loading with ExtendToNextSheet.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a CSV with a header and nine rows, loads it with TxtLoadOptions (MaxRowCount = 5, ExtendToNextSheet = true, HeaderRowsCount = 1) and verifies that two worksheets are generated, the first contains the header plus rows 1‑4, the second contains rows 5‑9, and that cells beyond the expected range remain empty.
    class Program
    {
        static void Main()
        {
            try
            {
                // Build CSV with 10 rows (header + 9 data rows) and two columns
                var sb = new StringBuilder();
                sb.AppendLine("Header1,Header2");
                for (int i = 1; i <= 9; i++)
                {
                    sb.AppendLine($"R{i},C{i}");
                }
                byte[] csvBytes = Encoding.UTF8.GetBytes(sb.ToString());

                // Load only 5 rows (including header) and enable overflow to next sheet
                TxtLoadOptions options = new TxtLoadOptions
                {
                    MaxRowCount = 5,
                    ExtendToNextSheet = true,
                    HeaderRowsCount = 1
                };

                // Load workbook from the CSV stream using the specified options
                Workbook workbook;
                using (MemoryStream ms = new MemoryStream(csvBytes))
                {
                    workbook = new Workbook(ms, options);
                }

                // Verify that a second worksheet was created for overflow rows
                if (workbook.Worksheets.Count != 2)
                    throw new InvalidOperationException($"Expected 2 worksheets, but found {workbook.Worksheets.Count}.");

                // First worksheet should contain header + first 4 data rows (R1‑R4)
                Worksheet sheet1 = workbook.Worksheets[0];
                ValidateCell(sheet1, 0, 0, "Header1");
                ValidateCell(sheet1, 1, 0, "R1");
                ValidateCell(sheet1, 4, 0, "R4");
                // Row index 5 should be empty in the first sheet
                if (!string.IsNullOrEmpty(sheet1.Cells[5, 0].StringValue))
                    throw new InvalidOperationException("Expected cell A6 in first sheet to be empty.");

                // Second worksheet should contain the remaining rows (R5‑R9)
                Worksheet sheet2 = workbook.Worksheets[1];
                ValidateCell(sheet2, 0, 0, "R5");
                ValidateCell(sheet2, 4, 0, "R9");
                // No extra rows beyond expected
                if (!string.IsNullOrEmpty(sheet2.Cells[5, 0].StringValue))
                    throw new InvalidOperationException("Expected cell A6 in second sheet to be empty.");

                Console.WriteLine("All checks passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to validate cell content
        private static void ValidateCell(Worksheet sheet, int row, int column, string expected)
        {
            string actual = sheet.Cells[row, column].StringValue;
            if (!string.Equals(actual, expected, StringComparison.Ordinal))
                throw new InvalidOperationException($"Cell ({row},{column}) expected '{expected}' but found '{actual}'.");
        }
    }
}
