// Title: C# Unit Test: Verify StyleFlag.QuotePrefix = false Does Not Reset Existing Quote Prefix in Aspose.Cells
// Description: Demonstrates a C# test that applies a QuotePrefix‑true style to a cell, then reapplies a style with StyleFlag.QuotePrefix disabled, confirming the original prefix remains unchanged and saving the workbook to a memory stream.
// Keywords: Aspose.Cells | StyleFlag | QuotePrefix | C# unit test | SetStyle | style flag testing | .NET | workbook | cell formatting | regression test
// Common Searches: Aspose.Cells StyleFlag QuotePrefix false unit test | preserve existing QuotePrefix when StyleFlag disabled | C# test SetStyle QuotePrefix flag behavior | how to verify QuotePrefix is not cleared in Aspose.Cells | unit testing cell style flags Aspose.Cells
// Developer Intent: Create an automated test that ensures a false StyleFlag.QuotePrefix does not alter a cell's current QuotePrefix setting.
// Use Cases: Validate that disabling the QuotePrefix flag leaves previously applied quote prefixes intact. | Prevent accidental loss of formatting when updating other style attributes on a cell. | Include in continuous‑integration pipelines to catch regressions related to style flag handling.
// AI Prompts: Generate an MSTest/NUnit/xUnit test method that asserts StyleFlag.QuotePrefix set to false preserves an existing QuotePrefix in Aspose.Cells for .NET. | Explain how StyleFlag influences SetStyle in Aspose.Cells and provide sample code for isolated unit testing of the QuotePrefix behavior. | Write a C# test that applies a style with QuotePrefix true, then reapplies a style with QuotePrefix false while the flag is disabled, verifies the property remains true, and saves the workbook to a MemoryStream.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Demonstrates a C# test that applies a QuotePrefix‑true style to a cell, then reapplies a style with StyleFlag.QuotePrefix disabled, confirming the original prefix remains unchanged and saving the workbook to a memory stream.
    public class StyleFlagQuotePrefixDemo
    {
        public static void Main()
        {
            try
            {
                QuotePrefixFlagFalseDoesNotClearExistingQuotePrefix();
                Console.WriteLine("Demo completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Demonstrates that when the StyleFlag.QuotePrefix is false,
        // the existing QuotePrefix setting on a cell is not cleared.
        private static void QuotePrefixFlagFalseDoesNotClearExistingQuotePrefix()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access cell A1 and put a simple numeric string
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("12345");

            // Create a style with QuotePrefix set to true
            Style style = workbook.CreateStyle();
            style.QuotePrefix = true;

            // Create a StyleFlag and enable the QuotePrefix flag
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true;

            // Apply the style to the cell using SetStyle with the flag
            cell.SetStyle(style, flag);

            // Verify that the QuotePrefix is now true
            if (!cell.GetStyle().QuotePrefix)
                throw new InvalidOperationException("QuotePrefix should be true after first application.");

            // Change the style to have QuotePrefix = false
            style.QuotePrefix = false;

            // Disable the QuotePrefix flag (set to false)
            flag.QuotePrefix = false;

            // Apply the style again; because the flag is false, the QuotePrefix setting should be ignored
            cell.SetStyle(style, flag);

            // Verify that the existing QuotePrefix value remains unchanged (still true)
            if (!cell.GetStyle().QuotePrefix)
                throw new InvalidOperationException("QuotePrefix should remain true when flag is false.");

            // Optional: Save to a memory stream to satisfy lifecycle rules (no file I/O needed)
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for potential further use
            }
        }
    }
}
