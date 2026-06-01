using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTests
{
    public class TxtLoadOptionsOverflowTests
    {
        public static void Main()
        {
            try
            {
                var test = new TxtLoadOptionsOverflowTests();
                test.OverflowRowsAreTransferredToSecondWorksheet();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public void OverflowRowsAreTransferredToSecondWorksheet()
        {
            // Prepare CSV data with 10 rows (including header) and two columns.
            var sb = new StringBuilder();
            sb.AppendLine("Header1,Header2"); // header row
            for (int i = 1; i <= 9; i++)
            {
                sb.AppendLine($"Data{i}A,Data{i}B");
            }
            string csvContent = sb.ToString();

            // Configure TxtLoadOptions:
            // - MaxRowCount = 5 (so only first 5 rows are loaded into the first sheet)
            // - ExtendToNextSheet = true (extra rows should go to a new sheet)
            var loadOptions = new TxtLoadOptions
            {
                MaxRowCount = 5,
                ExtendToNextSheet = true
                // HeaderRowsCount defaults to 1, which matches our CSV header.
            };

            // Load the CSV data into a workbook using a MemoryStream.
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(csvContent)))
            {
                var workbook = new Workbook(stream, loadOptions);

                // Verify that a second worksheet was created.
                AssertEqual(2, workbook.Worksheets.Count, "Expected a second worksheet due to overflow rows.");

                // First worksheet should contain exactly 5 rows (header + 4 data rows).
                var firstSheet = workbook.Worksheets[0];
                int firstSheetMaxRow = firstSheet.Cells.MaxRow; // zero‑based index of last row with data
                AssertEqual(4, firstSheetMaxRow, "First worksheet should have rows indexed 0‑4.");

                // Verify content of a known cell in the first sheet.
                AssertEqual("Header1", firstSheet.Cells[0, 0].StringValue, "Header cell mismatch.");
                AssertEqual("Data4A", firstSheet.Cells[4, 0].StringValue, "Data cell mismatch in first sheet.");

                // Second worksheet should contain the remaining rows (rows 5‑9 of the original CSV).
                var secondSheet = workbook.Worksheets[1];
                int secondSheetMaxRow = secondSheet.Cells.MaxRow;
                // There are 5 remaining rows (Data5‑Data9), so max row index should be 4.
                AssertEqual(4, secondSheetMaxRow, "Second worksheet should have rows indexed 0‑4.");

                // Verify that the first row of the second sheet contains the first overflow row.
                AssertEqual("Data5A", secondSheet.Cells[0, 0].StringValue, "First overflow row mismatch.");
                AssertEqual("Data9B", secondSheet.Cells[4, 1].StringValue, "Last overflow row mismatch.");

                // Optional: Save to a memory stream to ensure saving works without exceptions.
                using (var outStream = new MemoryStream())
                {
                    workbook.Save(outStream, SaveFormat.Xlsx);
                    if (outStream.Length == 0)
                        throw new InvalidOperationException("Workbook was not saved correctly.");
                }
            }
        }

        private void AssertEqual<T>(T expected, T actual, string message)
        {
            if (!object.Equals(expected, actual))
                throw new InvalidOperationException($"{message} Expected: {expected}, Actual: {actual}");
        }
    }
}