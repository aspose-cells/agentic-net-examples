// Title: Prevent Exponential Notation for Large Numbers in HTML with Aspose.Cells (.NET)
// Description: Demonstrates writing a very large integer to a worksheet and exporting the workbook to HTML using HtmlSaveOptions so the value appears as plain text rather than scientific notation.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportNumericFormat | PlainText | prevent exponential notation | large numbers HTML export | C# Aspose.Cells example | display big integers | remove scientific notation
// Common Searches: Aspose.Cells prevent scientific notation HTML | HtmlSaveOptions ExportNumericFormat plain text C# | export large integer to HTML Aspose.Cells | remove exponential format in HTML export Aspose | display 20‑digit number in HTML using Aspose.Cells
// Developer Intent: Export a workbook to HTML while keeping very large numeric values shown as plain text instead of scientific notation.
// Use Cases: Generate an HTML report that contains credit‑card numbers or other identifiers without losing digits. | Create financial statements where high‑precision figures must remain in full decimal form. | Build web‑based dashboards that display large IDs or timestamps exactly as entered.
// AI Prompts: Write C# code that sets HtmlSaveOptions.ExportNumericFormat to PlainText to stop exponential notation when exporting a workbook to HTML with Aspose.Cells. | Show an example that inserts a 20‑digit number into a cell and saves the worksheet as HTML, preserving the full digit sequence.

using System;
using Aspose.Cells;

// Demonstrates writing a very large integer to a worksheet and exporting the workbook to HTML using HtmlSaveOptions so the value appears as plain text rather than scientific notation.
class PreventExponentialNotation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a large numeric value as text to avoid exponential notation in HTML output
            sheet.Cells["A1"].PutValue("12345678901234567890");

            // Initialize HTML save options (default options are sufficient)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML using the configured options
            workbook.Save("LargeNumberPlainText.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
