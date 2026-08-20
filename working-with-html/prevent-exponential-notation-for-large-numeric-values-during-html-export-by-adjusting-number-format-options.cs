// Title: Avoid Scientific Notation for Large Numbers When Exporting to HTML Using Aspose.Cells (.NET)
// Description: C# sample that inserts a 20‑digit value, switches the workbook to 15‑digit SignificantDigitsType, applies a custom "0" number format, and saves to HTML with HtmlSaveOptions so the output shows the full integer instead of exponential form.
// Keywords: Aspose.Cells HTML export | C# large integer display | prevent exponential notation | SignificantDigitsType Digits15 | custom number format 0 | no scientific notation | Aspose.Cells .NET | HtmlSaveOptions | worksheet to HTML | large numeric values
// Common Searches: how to stop scientific notation in Aspose.Cells HTML output | Aspose.Cells export large numbers as plain text | set custom number format for HTML save in C# | SignificantDigitsType Digits15 usage example | prevent exponential format when converting workbook to HTML
// Developer Intent: The developer wants an HTML representation of a spreadsheet where very large numeric cells are rendered as full numbers, not in scientific notation.
// Use Cases: Publishing financial reports with 15‑digit account numbers that must appear unchanged in a web page. | Generating HTML invoices where product or SKU codes are long integers and should not be abbreviated. | Creating web‑ready dashboards that list identifiers or timestamps without exponent formatting.
// AI Prompts: Write C# code with Aspose.Cells that exports a worksheet to HTML while guaranteeing all numeric cells stay in plain integer format. | Explain the interaction between SignificantDigitsType and custom number formats during HTML conversion in Aspose.Cells. | Show how to apply a style to an entire column to keep large numbers from being displayed in scientific notation when saved as HTML.

using System;
using Aspose.Cells;

// C# sample that inserts a 20‑digit value, switches the workbook to 15‑digit SignificantDigitsType, applies a custom "0" number format, and saves to HTML with HtmlSaveOptions so the output shows the full integer instead of exponential form.
class PreventExponentialHtmlExport
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a large numeric value that would normally be rendered in exponential notation
        sheet.Cells["A1"].PutValue(12345678901234567890.0);

        // Set the global significant digits type to 15‑digit format to avoid scientific notation
        workbook.Settings.SignificantDigitsType = SignificantDigitsType.Digits15;

        // Apply a custom number format to the cell to force plain numeric display
        Style style = workbook.CreateStyle();
        style.Custom = "0";               // No decimal places, no exponent
        sheet.Cells["A1"].SetStyle(style);

        // Configure HTML save options (optional tweaks can be added here)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportFormula = false; // Export calculated values instead of formulas

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
