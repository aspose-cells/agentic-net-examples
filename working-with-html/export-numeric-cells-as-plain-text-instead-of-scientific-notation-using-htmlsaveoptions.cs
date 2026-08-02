// Title: Export numeric cells as plain text (no scientific notation) using Aspose.Cells HtmlSaveOptions in C#
// Description: Demonstrates how to create a workbook, apply a custom number format "0" to a large integer, configure HtmlSaveOptions with ExportDataOptions.All, and save the file as HTML so the numeric value appears as plain text instead of scientific notation.
// Keywords: Aspose.Cells | HtmlSaveOptions | C# | plain text numbers | prevent scientific notation | custom number format | HTML export | large integer | ExportDataOptions.All
// Common Searches: Aspose.Cells HTML export scientific notation | C# save workbook as HTML without exponent | force plain text numbers in Aspose.Cells HTML output | custom number format 0 Aspose.Cells
// Developer Intent: Generate an HTML file from a workbook where numeric cells are rendered as plain text rather than scientific notation.
// Use Cases: Financial reports that must display account numbers exactly as entered. | Web pages showing product serial numbers, IDs, or barcode values without exponent formatting. | Dashboard widgets that present large integer metrics with full precision.
// AI Prompts: Provide C# code that uses Aspose.Cells HtmlSaveOptions to keep custom number formats when exporting to HTML. | Show how to apply the "0" number format to a cell so the value is saved as plain text in HTML. | Explain the impact of HtmlSaveOptions.ExportDataOptions.All on style and format preservation during HTML export.

using System;
using Aspose.Cells;

namespace ExportNumericAsPlainText
{
    // Demonstrates how to create a workbook, apply a custom number format "0" to a large integer, configure HtmlSaveOptions with ExportDataOptions.All, and save the file as HTML so the numeric value appears as plain text instead of scientific notation.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Insert a large numeric value that would normally be shown in scientific notation
            // Example: 123456789012345
            Cell numericCell = cells["A1"];
            numericCell.PutValue(123456789012345);

            // Apply a custom number format to force plain text (no scientific notation)
            // The format "0" displays the full integer without exponent.
            Style style = workbook.CreateStyle();
            style.Custom = "0";
            numericCell.SetStyle(style);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Ensure all data (including styles) are exported
            htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

            // Save the workbook as HTML
            workbook.Save("NumericPlainText.html", htmlOptions);

            Console.WriteLine("Workbook saved as HTML with numeric cells displayed as plain text.");
        }
    }
}
