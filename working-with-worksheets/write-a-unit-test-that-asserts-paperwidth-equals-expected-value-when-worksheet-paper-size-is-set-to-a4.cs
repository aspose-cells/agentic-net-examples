// Title: C# unit test to verify worksheet PaperWidth equals A4 dimensions with Aspose.Cells
// Common Searches: aspocells c# unit test verify worksheet paper width for A4 | how to check worksheet PageSetup PaperWidth matches A4 size in Aspose.Cells | assert PaperWidth property equals expected A4 points in C# Aspose.Cells test | unit testing Aspose.Cells worksheet page setup paper size conversion to points
// Tags: Aspose.Cells PaperWidth verification | PageSetup PaperSize A4 verification | paper size to points conversion Aspose.Cells | page setup dimension tolerance | assert PaperWidth matches A4 Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a Workbook, sets the first worksheet's PageSetup PaperSize to A4, reads the PaperWidth in points, computes the expected width for A4 (210 mm converted to points), and throws an exception if the actual width differs from the expected value beyond a 0.5‑point tolerance, demonstrating a C# unit‑test pattern for Aspose.Cells.
    public class WorksheetPaperSizeTests
    {
        public static void Main(string[] args)
        {
            try
            {
                RunTest();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        private static void RunTest()
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Set the paper size of the worksheet to A4
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Retrieve the paper width (in points)
            double actualPaperWidth = sheet.PageSetup.PaperWidth;

            // Expected width for A4: 210 mm -> inches -> points
            double expectedPaperWidth = (210.0 / 25.4) * 72.0;

            // Verify within tolerance
            const double tolerance = 0.5;
            if (Math.Abs(expectedPaperWidth - actualPaperWidth) > tolerance)
            {
                throw new InvalidOperationException(
                    $"Paper width does not match A4 dimensions. Expected: {expectedPaperWidth}, Actual: {actualPaperWidth}");
            }
        }
    }
}
