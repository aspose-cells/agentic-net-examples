// Title: Write C# unit tests to confirm that StyleFlag.QuotePrefix = false preserves an existing quote‑prefix in Aspose.Cells cells
// AI Prompts: Generate a C# NUnit test that sets a cell value with a leading apostrophe, applies a style with StyleFlag.QuotePrefix set to false, and asserts the cell's QuotePrefix remains true. | Create a C# MSTest method that applies a style with StyleFlag.QuotePrefix true to a cell and verifies the QuotePrefix flag becomes true. | Provide a helper function that saves an Aspose.Cells workbook to a MemoryStream after style changes for further inspection.
// Common Searches: Aspose.Cells C# unit test for StyleFlag.QuotePrefix false behavior | How to keep leading apostrophe in Excel cell when applying style with Aspose.Cells | C# example verifying QuotePrefix flag is unchanged after SetStyle with flag disabled | Testing StyleFlag.QuotePrefix true to enable quote prefix in Aspose.Cells | Save Aspose.Cells workbook to MemoryStream in unit test
// Tags: Aspose.Cells StyleFlag QuotePrefix unit test | C# verify cell quote prefix persistence | SetStyle with StyleFlag false preserving quote prefix | Apply style with QuotePrefix true Aspose.Cells | MemoryStream workbook save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example defines two C# test methods that demonstrate how StyleFlag.QuotePrefix influences cell formatting in Aspose.Cells. One test confirms that applying a style with the flag set to false does not clear an existing quote‑prefix, while the other verifies that setting the flag to true correctly adds the quote‑prefix. A helper method shows how to save the workbook to a MemoryStream for additional validation.
    class Program
    {
        static void Main()
        {
            try
            {
                QuotePrefix_FlagFalse_DoesNotAlterExistingQuotePrefix();
                Console.WriteLine("QuotePrefix_FlagFalse_DoesNotAlterExistingQuotePrefix passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"QuotePrefix_FlagFalse_DoesNotAlterExistingQuotePrefix failed: {ex.Message}");
            }

            try
            {
                QuotePrefix_FlagTrue_AppliesStyleValue();
                Console.WriteLine("QuotePrefix_FlagTrue_AppliesStyleValue passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"QuotePrefix_FlagTrue_AppliesStyleValue failed: {ex.Message}");
            }
        }

        // Test: Setting StyleFlag.QuotePrefix to false should NOT remove an existing QuotePrefix.
        static void QuotePrefix_FlagFalse_DoesNotAlterExistingQuotePrefix()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a value that starts with a single quote (treated as text with QuotePrefix = true)
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("'12345");

            // Verify initial QuotePrefix is true
            if (!cell.GetStyle().QuotePrefix)
                throw new Exception("Initial QuotePrefix should be true after putting a leading quote.");

            // Create a new style with QuotePrefix explicitly set to false
            Style newStyle = workbook.CreateStyle();
            newStyle.QuotePrefix = false;

            // Create a StyleFlag where QuotePrefix flag is false (default)
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = false;

            // Apply the new style using the flag; existing QuotePrefix should remain unchanged
            cell.SetStyle(newStyle, flag);

            // Verify QuotePrefix is still true
            if (!cell.GetStyle().QuotePrefix)
                throw new Exception("QuotePrefix should remain true when StyleFlag.QuotePrefix is false.");
        }

        // Test: When the flag is true, the style's QuotePrefix value should be applied.
        static void QuotePrefix_FlagTrue_AppliesStyleValue()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Start with a cell without a leading quote
            Cell cell = sheet.Cells["B2"];
            cell.PutValue("67890");

            // Verify initial QuotePrefix is false
            if (cell.GetStyle().QuotePrefix)
                throw new Exception("Initial QuotePrefix should be false for a normal value.");

            // Create a style that sets QuotePrefix to true
            Style styleWithQuote = workbook.CreateStyle();
            styleWithQuote.QuotePrefix = true;

            // Create a flag that enables the QuotePrefix property
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true;

            // Apply the style with the flag
            cell.SetStyle(styleWithQuote, flag);

            // Verify QuotePrefix is now true
            if (!cell.GetStyle().QuotePrefix)
                throw new Exception("QuotePrefix should be true after applying style with flag enabled.");
        }

        // Helper method to save the workbook to a memory stream (demonstrates lifecycle usage)
        private static MemoryStream SaveWorkbookToStream(Workbook workbook)
        {
            MemoryStream stream = new MemoryStream();
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0;
            return stream;
        }
    }
}
