// Title: C# unit test to assert worksheet PaperWidth equals A4 size using Aspose.Cells PageSetup
// Common Searches: how to write a unit test for Aspose.Cells worksheet paper width A4 in C# | C# Aspose.Cells verify PaperWidth for A4 paper size | unit testing Aspose.Cells PageSetup PaperWidth against expected A4 dimensions | assert worksheet PageSetup PaperWidth matches 210 mm in inches using Aspose.Cells | Aspose.Cells C# test for correct A4 paper width tolerance
// Tags: Aspose.Cells worksheet PaperWidth verification | C# unit test Aspose.Cells PageSetup | A4 paper size width assertion Aspose.Cells | PaperWidth tolerance check Aspose.Cells | PageSetup PaperSizeType.PaperA4 validation

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a Workbook, accesses its first Worksheet, sets PageSetup.PaperSize to PaperA4, reads the PaperWidth in inches, computes the expected A4 width (210 mm ÷ 25.4), and throws an exception if the actual width differs from the expected value beyond a 0.001‑inch tolerance, demonstrating a C# unit test for Aspose.Cells.
    public class PaperSizeTests
    {
        public static void Main()
        {
            try
            {
                WorksheetPaperWidth_ShouldBe_A4Width();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public static void WorksheetPaperWidth_ShouldBe_A4Width()
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Set the paper size to A4
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Retrieve the paper width (in inches)
            double actualWidth = sheet.PageSetup.PaperWidth;

            // Expected A4 width: 210 mm converted to inches
            double expectedWidth = 210.0 / 25.4; // 1 inch = 25.4 mm

            // Verify the width matches the expected value within a small tolerance
            const double tolerance = 0.001;
            if (Math.Abs(expectedWidth - actualWidth) > tolerance)
            {
                throw new InvalidOperationException(
                    $"Paper width for A4 is incorrect. Expected: {expectedWidth}, Actual: {actualWidth}");
            }
        }
    }
}
