// Title: C# – Convert Excel to Compact HTML with Aspose.Cells PresentationPreference.AutoFit
// Description: Loads an .xlsx file using Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to true and WidthScalable to false, then saves the workbook as a tightly‑fitted HTML page with minimal column width waste.
// Keywords: Aspose.Cells | C# | Excel to HTML | PresentationPreference | AutoFit | compact column widths | HtmlSaveOptions | WidthScalable false | HTML export | web preview of spreadsheet
// Common Searches: Aspose.Cells export Excel to HTML auto‑fit columns | C# HtmlSaveOptions PresentationPreference example | How to disable scalable column widths in Aspose HTML export | compact HTML output from Excel using Aspose.Cells | generate web‑ready HTML from workbook C#
// Developer Intent: Generate an HTML representation of an Excel workbook where columns automatically shrink to fit content, producing a compact layout.
// Use Cases: Display spreadsheet previews on a website without horizontal scrolling. | Create email‑friendly HTML reports that preserve Excel column sizing. | Export dashboard worksheets for quick sharing while keeping the original layout.
// AI Prompts: Write C# code that opens an .xlsx file and saves it as HTML with PresentationPreference enabled and WidthScalable disabled using Aspose.Cells. | Explain the impact of HtmlSaveOptions.PresentationPreference on column widths in Aspose.Cells HTML export and list alternative settings for layout control. | Provide a step‑by‑step tutorial for configuring compact HTML output in a .NET console app with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Loads an .xlsx file using Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to true and WidthScalable to false, then saves the workbook as a tightly‑fitted HTML page with minimal column width waste.
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your file path)
            string excelPath = "input.xlsx";
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable presentation preference to get a more compact, auto‑fitted layout
            htmlOptions.PresentationPreference = true;

            // (Optional) Ensure column widths are not exported as scalable units
            // This keeps the columns tightly fitted as they appear in Excel.
            htmlOptions.WidthScalable = false;

            // Save the workbook as an HTML file
            string htmlPath = "output.html";
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Workbook successfully converted to HTML: {htmlPath}");
        }
    }
}
