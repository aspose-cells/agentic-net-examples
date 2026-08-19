// Title: Preserve a Leading Apostrophe in Excel Cells with Aspose.Cells .NET (StyleFlag QuotePrefix = false)
// Description: Demonstrates how to keep the initial single‑quote (apostrophe) in a cell when applying a style in Aspose.Cells for .NET. The example creates a workbook, inserts a value prefixed with an apostrophe, builds an empty Style, sets StyleFlag.QuotePrefix to false, applies the style, verifies the QuotePrefix remains true, and saves the file.
// Keywords: Aspose.Cells StyleFlag QuotePrefix false | C# preserve leading apostrophe Excel | Aspose.Cells keep cell apostrophe | SetStyle without removing QuotePrefix | Aspose.Cells .NET formatting apostrophe | Excel cell QuotePrefix property | Aspose.Cells example USA
// Common Searches: how to keep leading apostrophe when styling a cell Aspose.Cells | Aspose.Cells StyleFlag QuotePrefix false C# example | preserve single quote prefix after applying style in .NET | SetStyle does not change QuotePrefix Aspose.Cells | retain apostrophe in Excel cell after formatting
// Developer Intent: Apply a style to a cell while ensuring any existing apostrophe prefix remains unchanged.
// Use Cases: Formatting cells that contain text beginning with an apostrophe without stripping the prefix. | Generating Excel reports where data values include leading single quotes and still require visual styling. | Batch‑applying styles to ranges while preserving QuotePrefix flags for all affected cells.
// AI Prompts: Show C# code using Aspose.Cells StyleFlag to apply a style without altering the QuotePrefix property. | Give an example of formatting a range of cells while keeping any leading apostrophe prefixes intact. | Explain why setting StyleFlag.QuotePrefix to false preserves the apostrophe prefix during style application.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to keep the initial single‑quote (apostrophe) in a cell when applying a style in Aspose.Cells for .NET. The example creates a workbook, inserts a value prefixed with an apostrophe, builds an empty Style, sets StyleFlag.QuotePrefix to false, applies the style, verifies the QuotePrefix remains true, and saves the file.
    public class PreserveApostrophePrefixDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a value that starts with a single quote (apostrophe)
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue("'SampleText");

                // Create a style (no changes needed for QuotePrefix)
                Style style = workbook.CreateStyle();

                // Create a StyleFlag and explicitly set QuotePrefix to false
                // This ensures that applying the style will NOT modify the QuotePrefix property,
                // thus preserving the existing apostrophe prefix in the cell.
                StyleFlag flag = new StyleFlag
                {
                    QuotePrefix = false
                };

                // Apply the style to the cell using the StyleFlag
                cell.SetStyle(style, flag);

                // Verify that the QuotePrefix property remains true (since the cell value started with an apostrophe)
                Console.WriteLine("QuotePrefix after applying style: " + cell.GetStyle().QuotePrefix);

                // Save the workbook
                workbook.Save("PreserveApostrophePrefixDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            PreserveApostrophePrefixDemo.Run();
        }
    }
}
