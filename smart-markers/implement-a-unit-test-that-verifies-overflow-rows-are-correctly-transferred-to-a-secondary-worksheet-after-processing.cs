// Title: C# Unit Test: Verify TxtLoadOptions Overflow Rows Move to a Second Worksheet in Aspose.Cells
// Description: Shows how to load a CSV with TxtLoadOptions (MaxRowCount = 3, ExtendToNextSheet = true, HeaderRowsCount = 1), then asserts that a second worksheet is created and contains the rows that exceed the limit while the first sheet keeps the header and the initial data rows.
// Keywords: Aspose.Cells | TxtLoadOptions | ExtendToNextSheet | MaxRowCount | CSV to Excel | C# unit test | Workbook overflow rows | multiple worksheets | .NET Aspose.Cells example | split CSV sheets
// Common Searches: Aspose.Cells split CSV into multiple worksheets | TxtLoadOptions MaxRowCount example | ExtendToNextSheet property usage | C# unit test for overflow rows Aspose.Cells | How to move excess rows to a new sheet with Aspose.Cells
// Developer Intent: Confirm that rows beyond MaxRowCount are automatically placed on a new worksheet.
// Use Cases: Automated testing to ensure large CSV imports are divided across worksheets according to a row limit. | Validating that header rows stay on the first sheet while data rows overflow to subsequent sheets when ExtendToNextSheet is enabled. | Implementing a reliable CSV‑to‑Excel split routine in .NET applications using Aspose.Cells.
// AI Prompts: Generate an NUnit test that loads a CSV with TxtLoadOptions (MaxRowCount=3, ExtendToNextSheet=true) and verifies that overflow rows appear on a second worksheet. | Provide a C# snippet demonstrating how to configure TxtLoadOptions to limit rows per sheet and assert the workbook structure with Aspose.Cells. | Create a MSTest method that checks the header remains on the first sheet and remaining rows are transferred to additional worksheets when the row limit is exceeded.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Shows how to load a CSV with TxtLoadOptions (MaxRowCount = 3, ExtendToNextSheet = true, HeaderRowsCount = 1), then asserts that a second worksheet is created and contains the rows that exceed the limit while the first sheet keeps the header and the initial data rows.
    public class TxtLoadOptionsOverflowRowsTests
    {
        public static void Main()
        {
            try
            {
                new TxtLoadOptionsOverflowRowsTests().OverflowRowsAreTransferredToSecondWorksheet();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public void OverflowRowsAreTransferredToSecondWorksheet()
        {
            // Prepare CSV data with 5 rows (including header)
            string csvData =
                "Header1,Header2,Header3\n" +
                "R1C1,R1C2,R1C3\n" +
                "R2C1,R2C2,R2C3\n" +
                "R3C1,R3C2,R3C3\n" +
                "R4C1,R4C2,R4C3";

            // Configure TxtLoadOptions:
            // MaxRowCount = 3 (header + 2 data rows) -> rows beyond this go to next sheet
            // ExtendToNextSheet = true enables overflow handling
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                MaxRowCount = 3,
                ExtendToNextSheet = true,
                HeaderRowsCount = 1 // first row is header
            };

            // Load the CSV into a workbook using a memory stream
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
            {
                Workbook workbook = new Workbook(ms, loadOptions);

                // Verify that a second worksheet was created
                if (workbook.Worksheets.Count != 2)
                    throw new InvalidOperationException($"Expected 2 worksheets, found {workbook.Worksheets.Count}.");

                // First worksheet should contain header + first two data rows
                Worksheet sheet1 = workbook.Worksheets[0];
                Cells cells1 = sheet1.Cells;
                AssertEqual("Header1", cells1[0, 0].StringValue, "Header1");
                AssertEqual("R1C1", cells1[1, 0].StringValue, "R1C1");
                AssertEqual("R2C1", cells1[2, 0].StringValue, "R2C1");
                // No third data row in first sheet
                if (cells1[3, 0] != null && !string.IsNullOrEmpty(cells1[3, 0].StringValue))
                    throw new InvalidOperationException("Unexpected data in row 4 of first worksheet.");

                // Second worksheet should contain the remaining rows (R3 and R4)
                Worksheet sheet2 = workbook.Worksheets[1];
                Cells cells2 = sheet2.Cells;
                AssertEqual("R3C1", cells2[0, 0].StringValue, "R3C1");
                AssertEqual("R3C2", cells2[0, 1].StringValue, "R3C2");
                AssertEqual("R3C3", cells2[0, 2].StringValue, "R3C3");
                AssertEqual("R4C1", cells2[1, 0].StringValue, "R4C1");
                AssertEqual("R4C2", cells2[1, 1].StringValue, "R4C2");
                AssertEqual("R4C3", cells2[1, 2].StringValue, "R4C3");
            }
        }

        private void AssertEqual(string expected, string actual, string fieldName)
        {
            if (!string.Equals(expected, actual))
                throw new InvalidOperationException($"Assertion failed for {fieldName}: expected '{expected}', got '{actual}'.");
        }
    }
}
