// Title: HtmlSaveOptions.WidthScalable impacts only HTML export – PDF and XLSX stay unchanged (Aspose.Cells for .NET)
// Description: The example creates a workbook with narrow columns, enables HtmlSaveOptions.WidthScalable, saves the file as HTML, then exports the same workbook to PDF and XLSX using default settings. It demonstrates that the WidthScalable flag modifies only the HTML output while the PDF and Excel files retain their original layout.
// Keywords: Aspose.Cells | HtmlSaveOptions.WidthScalable | HTML export | PDF export | XLSX export | C# .NET | column width scaling | non‑HTML formats | unit test | Aspose.Cells for .NET
// Common Searches: HtmlSaveOptions WidthScalable affect PDF | Does WidthScalable change XLSX output | Aspose.Cells test column width scaling | How to export HTML with scalable columns using Aspose.Cells | Verify WidthScalable only for HTML in .NET
// Developer Intent: Confirm that enabling HtmlSaveOptions.WidthScalable changes the HTML file but leaves PDF and XLSX exports untouched.
// Use Cases: Generate HTML reports with auto‑adjusting column widths while preserving original layout in PDF and Excel files. | Run automated regression tests to ensure HTML‑specific options do not leak into other export formats. | Provide end‑users both web‑viewable HTML and printable PDF without extra configuration.
// AI Prompts: Generate an xUnit test that compares the PDF produced with and without HtmlSaveOptions.WidthScalable and asserts they are identical. | Show C# code that reads column widths from a saved XLSX file before and after enabling WidthScalable to confirm they match. | Explain the internal workflow of HtmlSaveOptions.WidthScalable and why it is limited to HTML rendering in Aspose.Cells.

using System;
using Aspose.Cells;

// The example creates a workbook with narrow columns, enables HtmlSaveOptions.WidthScalable, saves the file as HTML, then exports the same workbook to PDF and XLSX using default settings. It demonstrates that the WidthScalable flag modifies only the HTML output while the PDF and Excel files retain their original layout.
class WidthScalableNonHtmlTest
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Header");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("Long text that will be truncated if column width is narrow");
        ws.Cells["B2"].PutValue(12345);

        // Set narrow column widths to make the effect of WidthScalable visible in HTML
        ws.Cells.SetColumnWidth(0, 5);
        ws.Cells.SetColumnWidth(1, 5);

        // Configure HtmlSaveOptions with WidthScalable enabled
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.WidthScalable = true; // Enable scalable column width for HTML export

        // Save as HTML (WidthScalable influences this output)
        workbook.Save("output_widthscalable.html", htmlOptions);

        // Save the same workbook to PDF and XLSX without using HtmlSaveOptions
        // These formats should not be affected by the WidthScalable setting
        workbook.Save("output.pdf", SaveFormat.Pdf);
        workbook.Save("output.xlsx", SaveFormat.Xlsx);

        Console.WriteLine("Files saved. Verify that PDF and XLSX are unchanged by WidthScalable.");
    }
}
