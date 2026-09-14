// Title: Write an MSTest unit test in C# to verify that Aspose.Cells worksheet PageSetup.PaperWidth equals the A4 width (595 points)
// AI Prompts: Create an MSTest method that sets worksheet.PageSetup.PaperSize to PaperA4, reads PageSetup.PaperWidth, and asserts the value is within 0.5 points of 595. | Generate a NUnit test case for Aspose.Cells that configures a worksheet to A4 paper size and uses Assert.AreEqual with a tolerance to validate the PaperWidth property.
// Common Searches: how to assert worksheet paper width A4 using Aspose.Cells unit test | MSTest example for verifying Aspose.Cells PageSetup PaperWidth | C# unit test checking Aspose.Cells A4 paper dimensions | Aspose.Cells PageSetup PaperSize to A4 and validate PaperWidth in test
// Tags: Aspose.Cells worksheet paper width verification | C# test A4 paper size with PageSetup | MSTest Aspose.Cells PaperSize assertion | NUnit Aspose.Cells PaperWidth tolerance check | PageSetup A4 size verification

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates an in‑memory Workbook, sets the first worksheet's PageSetup.PaperSize to A4, reads the read‑only PageSetup.PaperWidth, and asserts that the width is within 0.5 points of the expected 595‑point A4 width, demonstrating how to write a unit test for this validation.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (in-memory, no file I/O)
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Set the paper size to A4 (standard size)
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                // Retrieve the actual paper width from the worksheet's PageSetup (read‑only)
                double actualPaperWidth = worksheet.PageSetup.PaperWidth;

                // Expected A4 width in points (1 point = 1/72 inch, A4 width = 8.27 inches ≈ 595 points)
                const double expectedA4Width = 595.0;

                // Verify that the actual width is within a small tolerance of the expected value
                const double tolerance = 0.5; // points
                if (Math.Abs(expectedA4Width - actualPaperWidth) <= tolerance)
                {
                    Console.WriteLine($"Success: PaperWidth is {actualPaperWidth} points, within tolerance of {tolerance} points.");
                }
                else
                {
                    Console.WriteLine($"Failure: Expected PaperWidth {expectedA4Width} points, but got {actualPaperWidth} points.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display them
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
