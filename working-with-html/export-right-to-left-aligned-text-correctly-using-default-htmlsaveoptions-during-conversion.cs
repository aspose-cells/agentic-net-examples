// Title: Export RTL Text to HTML with Default HtmlSaveOptions – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable the DisplayRightToLeft flag on a worksheet, insert Arabic text, and save the workbook as HTML using the out‑of‑the‑box HtmlSaveOptions, preserving right‑to‑left alignment without extra configuration.
// Keywords: Aspose.Cells HTML export RTL | DisplayRightToLeft C# | default HtmlSaveOptions | Arabic Excel to HTML | right‑to‑left alignment Aspose.Cells | C# Excel to HTML conversion
// Common Searches: Aspose.Cells save HTML RTL text | How to keep right‑to‑left direction when exporting Excel to HTML | DisplayRightToLeft property HTML output example | Export Arabic worksheet to HTML using Aspose.Cells .NET | Default HtmlSaveOptions RTL support
// Developer Intent: Generate an HTML file from a workbook that contains right‑to‑left language text while retaining proper text direction using the default save options.
// Use Cases: Create web‑ready reports for Arabic, Hebrew, or other RTL languages without custom CSS. | Automate batch conversion of Excel files to HTML for multilingual portals. | Produce printable HTML pages that maintain the original worksheet layout and directionality.
// AI Prompts: Write C# code that saves a workbook with Arabic text to HTML using Aspose.Cells and explain the role of DisplayRightToLeft. | Show how to validate that the exported HTML renders RTL correctly and list the HTML attributes added by default HtmlSaveOptions. | Suggest a way to extend the default HtmlSaveOptions to inject a custom CSS class for RTL cells while keeping built‑in direction handling.

using System;
using Aspose.Cells;

namespace AsposeCellsRtlHtmlExport
{
    // Demonstrates how to enable the DisplayRightToLeft flag on a worksheet, insert Arabic text, and save the workbook as HTML using the out‑of‑the‑box HtmlSaveOptions, preserving right‑to‑left alignment without extra configuration.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Enable right‑to‑left display for the worksheet
            sheet.DisplayRightToLeft = true;

            // Put some right‑to‑left text (Arabic example)
            sheet.Cells["A1"].PutValue("نص من اليمين إلى اليسار");

            // Use the default HtmlSaveOptions (no custom options required)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML (default save rule)
            workbook.Save("RtlExport.html", htmlOptions);

            Console.WriteLine("Workbook saved to RtlExport.html with right‑to‑left alignment.");
        }
    }
}
