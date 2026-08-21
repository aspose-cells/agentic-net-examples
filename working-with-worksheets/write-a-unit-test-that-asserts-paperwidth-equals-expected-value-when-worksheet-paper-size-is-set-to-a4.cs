// Title: C# unit test for Aspose.Cells worksheet PaperWidth with A4 paper size
// Description: Creates a workbook, sets the first worksheet's PageSetup.PaperSize to PaperA4, reads PageSetup.PaperWidth, converts 210 mm to inches, and asserts the value matches the expected width within a 0.0001‑inch tolerance using a .NET testing framework.
// Keywords: Aspose.Cells | C# | unit test | PaperWidth | A4 | PageSetup | PaperSizeType | worksheet | expected width | inches | tolerance | MSTest | xUnit | NUnit
// Common Searches: Aspose.Cells unit test PaperWidth A4 | assert worksheet paper width C# | validate A4 dimensions Aspose.Cells | C# test PageSetup PaperSize | PaperWidth tolerance check Aspose.Cells
// Developer Intent: Write a test that confirms the worksheet's PaperWidth equals the A4 width (210 mm) converted to inches, allowing a small tolerance.
// Use Cases: Verify printed or exported reports use correct A4 dimensions before release. | Add page‑setup validation to continuous‑integration pipelines. | Detect regressions in paper‑size handling after upgrading Aspose.Cells. | Automate compliance checks for printing standards in enterprise applications. | Provide a reference for test‑driven development with Aspose.Cells.
// AI Prompts: Generate an MSTest method that asserts sheet.PageSetup.PaperWidth equals 210 mm in inches with a 0.0001‑inch tolerance. | Create an xUnit test for Aspose.Cells verifying PaperWidth after setting PaperSizeType.PaperA4. | Write a NUnit test that checks the PaperWidth property matches the expected A4 width when PaperSize is set to PaperA4.

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a workbook, sets the first worksheet's PageSetup.PaperSize to PaperA4, reads PageSetup.PaperWidth, converts 210 mm to inches, and asserts the value matches the expected width within a 0.0001‑inch tolerance using a .NET testing framework.
    public class PageSetupDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set the paper size to A4
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                // Retrieve the paper width in inches
                double actualWidth = sheet.PageSetup.PaperWidth;

                // Expected width for A4 (210 mm) converted to inches
                double expectedWidth = 210.0 / 25.4; // ≈ 8.26771653543307

                // Tolerance for comparison
                double tolerance = 0.0001;

                // Verify the width matches the expected value within tolerance
                if (Math.Abs(expectedWidth - actualWidth) <= tolerance)
                {
                    Console.WriteLine("Paper width for A4 matches expected value.");
                }
                else
                {
                    Console.WriteLine($"Paper width mismatch. Expected: {expectedWidth}, Actual: {actualWidth}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
