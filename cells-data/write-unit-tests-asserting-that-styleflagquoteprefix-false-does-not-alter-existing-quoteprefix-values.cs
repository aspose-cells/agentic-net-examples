// Title: C# unit test to verify that StyleFlag.QuotePrefix set to false does not change an existing true QuotePrefix in Aspose.Cells
// Common Searches: Aspose.Cells unit test for preserving QuotePrefix when StyleFlag.QuotePrefix is false | C# test that SetStyle with QuotePrefix flag disabled keeps original quote prefix | How to assert QuotePrefix remains true after applying style with disabled QuotePrefix flag in Aspose.Cells | Write NUnit test for StyleFlag QuotePrefix immutability in Aspose.Cells | Verify that StyleFlag false does not overwrite cell QuotePrefix in Aspose.Cells C#
// Tags: Aspose.Cells StyleFlag QuotePrefix immutability | C# SetStyle flag false behavior Aspose.Cells | unit test preserving cell quote prefix Aspose.Cells | NUnit Aspose.Cells style flag verification | cell style flag handling Aspose.Cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, sets cell A1's QuotePrefix to true using a StyleFlag with QuotePrefix enabled, then attempts to set QuotePrefix to false with a StyleFlag where QuotePrefix is disabled, and includes unit test assertions that the original true value remains unchanged.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                var cell = worksheet.Cells["A1"];
                cell.PutValue("Sample");

                // Apply QuotePrefix = true using a StyleFlag that enables the property
                var style = cell.GetStyle();
                style.QuotePrefix = true;
                var flag = new StyleFlag();
                flag.QuotePrefix = true;
                cell.SetStyle(style, flag);

                // Verify that the QuotePrefix flag is set
                if (!cell.GetStyle().QuotePrefix)
                    throw new Exception("QuotePrefix should be true after initial setting.");

                // Attempt to change QuotePrefix to false using a StyleFlag that disables the property
                var newStyle = cell.GetStyle();
                newStyle.QuotePrefix = false; // This value will be ignored because the flag is false
                var newFlag = new StyleFlag();
                newFlag.QuotePrefix = false;
                cell.SetStyle(newStyle, newFlag);

                // Verify that the original QuotePrefix value remains unchanged
                if (!cell.GetStyle().QuotePrefix)
                    throw new Exception("QuotePrefix should remain true when the StyleFlag disables modification.");

                Console.WriteLine("QuotePrefix tests passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
