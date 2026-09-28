// Title: Batch enable QuotePrefix for every cell in a worksheet with StyleFlag in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a style with QuotePrefix set to true, configures a StyleFlag to affect only QuotePrefix, and applies the style to all cells of a worksheet using Aspose.Cells. | Show how to use Cells.ApplyStyle together with a StyleFlag to add the leading apostrophe to an entire sheet without altering other cell formatting. | Demonstrate saving the workbook after batch updating QuotePrefix to verify the changes.
// Common Searches: asp.net aspose.cells batch set QuotePrefix for whole worksheet | c# apply leading apostrophe to all cells using StyleFlag | how to use ApplyStyle with StyleFlag to change only QuotePrefix in Aspose.Cells | bulk update cell style property QuotePrefix Aspose.Cells .NET | set quote prefix for every cell programmatically Aspose.Cells C#
// Tags: Aspose.Cells ApplyStyle QuotePrefix | StyleFlag selective formatting C# | batch cell style update Aspose.Cells | leading apostrophe worksheet .NET | bulk QuotePrefix change Aspose.Cells | cells.ApplyStyle with StyleFlag

using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixBatchUpdate
{
    // The example creates a workbook, adds sample data, defines a style with QuotePrefix enabled, sets a StyleFlag to apply only the QuotePrefix property, and uses Cells.ApplyStyle to batch‑apply this style to every cell in the first worksheet. The workbook is then saved as QuotePrefixBatchUpdated.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data (including values that start with a single quote)
            cells["A1"].PutValue("'TextWithQuote");
            cells["A2"].PutValue("NormalText");
            cells["B1"].PutValue("'12345");
            cells["B2"].PutValue(9876);

            // Create a style and enable the QuotePrefix property
            Style quoteStyle = workbook.CreateStyle();
            quoteStyle.QuotePrefix = true; // This indicates the cell should have QuotePrefix applied

            // Create a StyleFlag and enable the QuotePrefix flag
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true; // Only the QuotePrefix setting will be applied

            // Apply the style to the entire worksheet using the flag
            // This batches the QuotePrefix update for every cell in the sheet
            cells.ApplyStyle(quoteStyle, flag);

            // Save the workbook to verify the changes
            workbook.Save("QuotePrefixBatchUpdated.xlsx");
        }
    }
}
