// Title: Convert Excel to Minimal HTML (Disable CSS) with Aspose.Cells for .NET
// Description: Loads an .xlsx file, sets HtmlSaveOptions.DisableCss to true, and saves the workbook as lightweight HTML that contains only inline styling using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | Excel to HTML | DisableCss | HtmlSaveOptions | minimal HTML | inline styles | export workbook | no CSS | .NET HTML conversion
// Common Searches: Aspose.Cells export Excel to HTML without CSS | C# HtmlSaveOptions DisableCss example | convert workbook to plain HTML Aspose | generate lightweight HTML from Excel .NET | save Excel as HTML inline styles only
// Developer Intent: Produce an HTML file from an Excel workbook while suppressing external CSS, yielding a compact document with inline formatting.
// Use Cases: Publish Excel data on web pages where external style sheets are prohibited. | Embed spreadsheet content in email messages that require self‑contained HTML. | Archive workbook snapshots as plain HTML to reduce storage overhead.
// AI Prompts: Show how to convert an Excel workbook to HTML with Aspose.Cells and turn off CSS generation. | Provide a C# snippet that saves an .xlsx file as minimal HTML using HtmlSaveOptions.DisableCss. | Explain the steps to configure Aspose.Cells HtmlSaveOptions for CSS‑free HTML output and handle embedded images.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlConversion
{
    // Loads an .xlsx file, sets HtmlSaveOptions.DisableCss to true, and saves the workbook as lightweight HTML that contains only inline styling using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path for the generated HTML file
            string htmlPath = "output.html";

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable CSS generation to produce minimal inline‑styled HTML
            htmlOptions.DisableCss = true;

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Workbook converted to HTML at: {htmlPath}");
        }
    }
}
