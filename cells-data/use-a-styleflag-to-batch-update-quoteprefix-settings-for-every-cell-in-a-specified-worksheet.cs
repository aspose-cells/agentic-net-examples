// Title: Apply QuotePrefix to Every Cell in a Worksheet with Aspose.Cells .NET using StyleFlag
// Description: Demonstrates how to create a style with QuotePrefix enabled, configure a StyleFlag to target only that property, and batch‑apply the style to all cells of a worksheet in a single call. The example preserves existing formatting, verifies the setting, and saves the workbook.
// Keywords: Aspose.Cells | .NET | C# | StyleFlag | QuotePrefix | ApplyStyle | batch update cell style | single quote prefix | Excel worksheet | preserve leading apostrophe
// Common Searches: Aspose.Cells set QuotePrefix for whole sheet | StyleFlag apply only QuotePrefix .NET | batch apply QuotePrefix to all cells | preserve leading single quote in Excel using Aspose | ApplyStyle with StyleFlag example
// Developer Intent: Enable the QuotePrefix flag on every cell of a worksheet in one operation while leaving other style attributes untouched.
// Use Cases: Import CSV files where values start with a single quote and must remain text in Excel. | Generate reports that require all cells to display literal strings without automatic number conversion. | Prepare workbooks for export where each cell’s original text representation must be retained.
// AI Prompts: Write C# code that uses Aspose.Cells to set QuotePrefix for an entire worksheet with StyleFlag. | Explain how StyleFlag works with ApplyStyle to modify only the QuotePrefix property. | Show how to load an existing workbook and batch update QuotePrefix for all cells while keeping existing formatting.

using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixBatchUpdate
{
    // Demonstrates how to create a style with QuotePrefix enabled, configure a StyleFlag to target only that property, and batch‑apply the style to all cells of a worksheet in a single call. The example preserves existing formatting, verifies the setting, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // new workbook with one default worksheet

            // Get reference to the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that includes values starting with a single quote
            sheet.Cells["A1"].PutValue("'TextWithQuote");
            sheet.Cells["B2"].PutValue("'12345");
            sheet.Cells["C3"].PutValue("NormalText");

            // Create a style and enable the QuotePrefix property
            Style quoteStyle = workbook.CreateStyle();
            quoteStyle.QuotePrefix = true; // this indicates that the cell's value starts with a single quote

            // Create a StyleFlag and enable only the QuotePrefix flag
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true; // apply only the QuotePrefix setting

            // Apply the style to all cells in the worksheet using the flag
            // This batches the update, so every cell will have QuotePrefix set to true
            sheet.Cells.ApplyStyle(quoteStyle, flag);

            // Verify the result for a few cells
            Console.WriteLine("A1 QuotePrefix: " + sheet.Cells["A1"].GetStyle().QuotePrefix);
            Console.WriteLine("B2 QuotePrefix: " + sheet.Cells["B2"].GetStyle().QuotePrefix);
            Console.WriteLine("C3 QuotePrefix: " + sheet.Cells["C3"].GetStyle().QuotePrefix);

            // Save the workbook to a file
            workbook.Save("QuotePrefixBatchUpdated.xlsx");
        }
    }
}
