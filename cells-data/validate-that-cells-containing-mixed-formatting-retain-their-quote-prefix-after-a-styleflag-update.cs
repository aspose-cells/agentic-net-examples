// Title: How to keep a cell's QuotePrefix when updating only background color with a StyleFlag in Aspose.Cells for .NET
// AI Prompts: Show C# code that updates a cell's background color using StyleFlag while preserving its existing QuotePrefix in Aspose.Cells. | Generate a minimal Aspose.Cells example that applies a new style with the CellShading flag only and confirms the QuotePrefix flag remains set.
// Common Searches: Aspose.Cells preserve leading apostrophe after applying StyleFlag in C# | QuotePrefix flag lost when changing cell shading with StyleFlag Aspose.Cells .NET | How to update cell background without resetting QuotePrefix using Aspose.Cells API
// Tags: QuotePrefix preservation with StyleFlag | Aspose.Cells selective style flag background | C# update cell shading without clearing quote prefix | Aspose.Cells style flag cell formatting | Excel leading single quote retention Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixValidation
{
    // The example creates a workbook, writes a value with a leading single quote, applies an initial style that enables QuotePrefix, then changes only the background color and font boldness using a StyleFlag. It verifies that the QuotePrefix flag stays true while the new formatting is applied.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access cell A1 and put a text value that starts with a single quote
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("'MixedFormatting");

            // Create an initial style with QuotePrefix enabled and some formatting (e.g., bold, red font)
            Style initialStyle = workbook.CreateStyle();
            initialStyle.QuotePrefix = true;               // Enable quote prefix
            initialStyle.Font.IsBold = true;               // Apply bold formatting
            initialStyle.Font.Color = System.Drawing.Color.Red; // Apply font color

            // Apply the initial style to the cell
            cell.SetStyle(initialStyle);

            // Verify that QuotePrefix is set initially
            Console.WriteLine("Initial QuotePrefix: " + cell.GetStyle().QuotePrefix); // Expected: True

            // Create a new style that changes only the background color
            Style newStyle = workbook.CreateStyle();
            newStyle.ForegroundColor = System.Drawing.Color.Yellow;
            newStyle.Pattern = BackgroundType.Solid;

            // Create a StyleFlag that applies only the cell shading (background) and font changes,
            // but does NOT include the QuotePrefix flag
            StyleFlag flag = new StyleFlag
            {
                CellShading = true,
                FontBold = true   // Example of another flag; QuotePrefix remains false
            };

            // Apply the new style with the flag to the same cell
            cell.SetStyle(newStyle, flag);

            // After applying the style flag, verify that QuotePrefix is still retained
            bool quotePrefixAfterUpdate = cell.GetStyle().QuotePrefix;
            Console.WriteLine("QuotePrefix after StyleFlag update: " + quotePrefixAfterUpdate); // Expected: True

            // Additional verification: other formatting changes should be applied
            Style resultingStyle = cell.GetStyle();
            Console.WriteLine("Background color applied: " + (resultingStyle.Pattern == BackgroundType.Solid));
            Console.WriteLine("Foreground color applied: " + (resultingStyle.ForegroundColor.ToArgb() == System.Drawing.Color.Yellow.ToArgb()));

            // Save the workbook (optional, just to complete lifecycle)
            workbook.Save("QuotePrefixValidation.xlsx");
        }
    }
}
