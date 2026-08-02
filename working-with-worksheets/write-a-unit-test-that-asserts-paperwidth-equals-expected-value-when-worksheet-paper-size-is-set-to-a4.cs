// Title: C# unit test for Aspose.Cells Worksheet PaperWidth with A4 paper size
// Description: Shows how to write a .NET unit test that sets a worksheet's PageSetup.PaperSize to PaperA4, reads the PageSetup.PaperWidth in inches, computes the expected width from 210 mm (210/25.4), and asserts the value within a 0.001‑inch tolerance using MSTest, NUnit, or xUnit.
// Keywords: Aspose.Cells | C# | .NET | Worksheet | PageSetup | PaperWidth | PaperSize | PaperA4 | PaperSizeType | unit test | MSTest | NUnit | xUnit | assertion | tolerance | print layout validation
// Common Searches: Aspose.Cells unit test PaperWidth | assert worksheet PaperWidth A4 C# | PageSetup PaperSize A4 test example | verify PaperWidth inches Aspose.Cells | C# test for PaperSizeType.PaperA4
// Developer Intent: Write a unit test that confirms the worksheet PaperWidth matches the expected A4 width.
// Use Cases: Validate that setting PaperSize to PaperA4 yields the correct printable width for PDF or print output. | Include regression checks in CI pipelines to catch changes in page‑setup calculations. | Ensure consistent page dimensions when exporting workbooks to formats that rely on physical size.
// AI Prompts: Generate an MSTest method that creates a Workbook, sets the first worksheet's PaperSize to PaperA4, and asserts PaperWidth equals 210/25.4 inches with a 0.001 tolerance. | Provide an NUnit test case for Aspose.Cells that verifies PaperWidth after assigning PaperSizeType.PaperA4, including necessary using statements and Assert statements. | Write a xUnit test for Aspose.Cells .NET that checks the PaperWidth of a worksheet set to A4 and fails with a clear message if the difference exceeds 0.001 inches.

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Shows how to write a .NET unit test that sets a worksheet's PageSetup.PaperSize to PaperA4, reads the PageSetup.PaperWidth in inches, computes the expected width from 210 mm (210/25.4), and asserts the value within a 0.001‑inch tolerance using MSTest, NUnit, or xUnit.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set the paper size to A4 (210 mm x 297 mm)
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                // Retrieve the paper width in inches
                double actualWidth = sheet.PageSetup.PaperWidth;

                // Expected width for A4: 210 mm converted to inches
                double expectedWidth = 210.0 / 25.4; // ≈ 8.2677 inches

                // Define a tolerance for floating‑point comparison
                double tolerance = 0.001;

                // Verify that the actual width matches the expected width within the tolerance
                if (Math.Abs(expectedWidth - actualWidth) <= tolerance)
                {
                    Console.WriteLine($"Success: PaperWidth is {actualWidth:F4} inches, matches expected {expectedWidth:F4} inches.");
                }
                else
                {
                    Console.WriteLine($"Failure: PaperWidth is {actualWidth:F4} inches, expected {expectedWidth:F4} inches.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
