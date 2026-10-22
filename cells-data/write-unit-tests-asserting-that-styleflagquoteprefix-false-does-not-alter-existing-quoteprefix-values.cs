// Title: C# unit test to verify that resetting StyleFlag.QuotePrefix to false leaves the cell value unchanged in Aspose.Cells
// Common Searches: how to assert that StyleFlag.QuotePrefix false does not change cell data in Aspose.Cells C# unit test | Aspose.Cells C# test for QuotePrefix flag reset without modifying cell value | unit testing QuotePrefix property behavior in Aspose.Cells workbook using .NET | verify that disabling QuotePrefix in Aspose.Cells does not alter existing cell content | C# example of checking QuotePrefix flag impact on cell value in Aspose.Cells
// Tags: Aspose.Cells StyleFlag QuotePrefix reset test | C# unit testing Aspose.Cells cell style flag | verify cell value stability after QuotePrefix change | Aspose.Cells workbook cell style flag behavior | C# .NET Aspose.Cells quote prefix validation

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a workbook, sets QuotePrefix = true on a cell, confirms the flag, captures the cell's value, then resets QuotePrefix to false, asserts the flag is false, and verifies that the original cell value remains unchanged.
    public class QuotePrefixTests
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cell = sheet.Cells["A1"];

                // Set an initial value
                cell.PutValue("TestValue");

                // Apply QuotePrefix = true
                var style = cell.GetStyle();
                style.QuotePrefix = true;
                cell.SetStyle(style);

                // Verify the flag is true
                if (!cell.GetStyle().QuotePrefix)
                    throw new Exception("QuotePrefix should be true after setting it.");

                // Capture the current cell value
                var originalValue = cell.Value;

                // Reset QuotePrefix = false
                var style2 = cell.GetStyle();
                style2.QuotePrefix = false;
                cell.SetStyle(style2);

                // Verify the flag is false
                if (cell.GetStyle().QuotePrefix)
                    throw new Exception("QuotePrefix should be false after resetting it.");

                // Ensure the cell's value has not changed
                if (!object.Equals(originalValue, cell.Value))
                    throw new Exception("Cell value must remain unchanged when QuotePrefix is set to false.");

                Console.WriteLine("QuotePrefix test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }
}
