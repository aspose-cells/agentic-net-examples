using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    class PageSetupTests
    {
        static void Main(string[] args)
        {
            try
            {
                // Execute the test logic
                VerifyPaperWidthForA4();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                // Report any unexpected errors
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Verifies that the PaperWidth property for an A4 page size matches the expected value (~8.27 inches).
        /// </summary>
        private static void VerifyPaperWidthForA4()
        {
            // Create a new workbook (a default worksheet is added automatically)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the paper size to A4
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Retrieve the paper width in inches
            double actualPaperWidth = worksheet.PageSetup.PaperWidth;

            // Expected width for A4 paper (210 mm) in inches ≈ 8.27 inches
            const double expectedPaperWidth = 8.27;
            const double tolerance = 0.01;

            // Validate the result within the tolerance
            if (Math.Abs(expectedPaperWidth - actualPaperWidth) > tolerance)
            {
                throw new InvalidOperationException(
                    $"PaperWidth for A4 should be approximately {expectedPaperWidth} inches, but was {actualPaperWidth} inches.");
            }
        }
    }
}