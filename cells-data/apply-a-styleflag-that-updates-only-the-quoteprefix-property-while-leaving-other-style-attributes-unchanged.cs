// Title: Set only the QuotePrefix flag with StyleFlag in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use Aspose.Cells' StyleFlag to modify the QuotePrefix property of a cell without affecting any other formatting. The example creates a workbook, applies a style that toggles QuotePrefix on cell A1, verifies the change, and saves the file.
// Keywords: Aspose.Cells | StyleFlag | QuotePrefix | C# | .NET | cell formatting | update single style attribute | Excel export | programmatic Excel styling
// Common Searches: Aspose.Cells change only QuotePrefix | StyleFlag example C# | set cell as text without altering format Aspose.Cells | how to toggle QuotePrefix programmatically | apply single style property Aspose.Cells
// Developer Intent: Change the QuotePrefix attribute of a specific cell while keeping all existing style settings intact.
// Use Cases: Force Excel to treat numeric strings as text in a single cell without modifying font, color, or borders. | Batch‑toggle QuotePrefix across a range while preserving each cell's original number format and styling. | Integrate QuotePrefix handling into automated report generation where existing cell styles must remain unchanged.
// AI Prompts: Generate C# code that sets QuotePrefix = true for cell B2 using Aspose.Cells StyleFlag while preserving its current style. | Show an example that flips the QuotePrefix flag on a cell, verifies the property, and then reverts it, using StyleFlag. | Explain the purpose of StyleFlag in Aspose.Cells and illustrate how to apply it to a range to toggle QuotePrefix without affecting other formatting.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to use Aspose.Cells' StyleFlag to modify the QuotePrefix property of a cell without affecting any other formatting. The example creates a workbook, applies a style that toggles QuotePrefix on cell A1, verifies the change, and saves the file.
    public class QuotePrefixStyleFlagDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access cell A1 and put a value that looks like a number
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue("123456");

                // Create a style and enable the QuotePrefix property
                Style style = workbook.CreateStyle();
                style.QuotePrefix = true; // Desired value for the cell

                // Create a StyleFlag and enable only the QuotePrefix flag
                StyleFlag flag = new StyleFlag
                {
                    QuotePrefix = true   // All other flags remain false
                };

                // Apply the style to the cell using the flag – only QuotePrefix will be updated
                cell.SetStyle(style, flag);

                // Verify the applied QuotePrefix value (optional)
                Console.WriteLine("QuotePrefix after first apply: " + cell.GetStyle().QuotePrefix);

                // Change QuotePrefix to false while keeping other style attributes unchanged
                style.QuotePrefix = false;
                cell.SetStyle(style, flag);

                // Verify the updated QuotePrefix value (optional)
                Console.WriteLine("QuotePrefix after second apply: " + cell.GetStyle().QuotePrefix);

                // Save the workbook
                string outputPath = "QuotePrefixStyleFlagDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            QuotePrefixStyleFlagDemo.Run();
        }
    }
}
