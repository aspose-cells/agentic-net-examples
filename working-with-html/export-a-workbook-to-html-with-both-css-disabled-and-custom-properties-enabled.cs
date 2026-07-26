// Title: Export Workbook to HTML with Inline Styles and CSS Custom Properties (C# Aspose.Cells)
// Description: Creates a workbook, adds sample data, configures HtmlSaveOptions to disable external CSS (inline styles only) and enable CSS custom properties, then saves the file as HTML using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableCss | EnableCssCustomProperties | C# | Excel to HTML | inline CSS | custom properties | HTML export | lightweight HTML
// Common Searches: Aspose.Cells export HTML without external CSS | HtmlSaveOptions DisableCss true example | EnableCssCustomProperties C# Aspose.Cells | Generate HTML from Excel with inline styles | How to reduce HTML size when exporting Excel
// Developer Intent: Generate an HTML representation of an Excel workbook using Aspose.Cells while suppressing external style sheets and activating CSS custom‑property optimization.
// Use Cases: Embedding Excel data in web pages where loading extra CSS files is undesirable. | Creating HTML email templates that require inline styling for consistent client rendering. | Producing compact HTML previews of large workbooks for fast loading in web applications.
// AI Prompts: Write C# code that loads an existing .xlsx file and saves it as HTML with DisableCss=true and EnableCssCustomProperties=true using Aspose.Cells. | Explain the benefits of EnableCssCustomProperties in the HTML output and scenarios where it improves performance. | Show how to adjust HtmlSaveOptions to embed worksheet images as Base64 strings while keeping CSS disabled.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Creates a workbook, adds sample data, configures HtmlSaveOptions to disable external CSS (inline styles only) and enable CSS custom properties, then saves the file as HTML using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B2"].PutValue(12345);

            // Configure HTML save options:
            // - DisableCss = true  => use only inline styles, no external CSS.
            // - EnableCssCustomProperties = true => enable CSS custom properties optimization.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                DisableCss = true,
                EnableCssCustomProperties = true
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save("ExportedWorkbook.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML with CSS disabled and custom properties enabled.");
        }
    }
}
