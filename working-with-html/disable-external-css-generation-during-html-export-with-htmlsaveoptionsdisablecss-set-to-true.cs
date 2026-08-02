// Title: Export Excel to HTML with inline CSS only using Aspose.Cells HtmlSaveOptions.DisableCss (C#)
// Description: A C# example that creates a workbook, applies cell formatting, sets HtmlSaveOptions.DisableCss = true, and saves the file as a self‑contained HTML page with all styles embedded inline—no external CSS file is generated.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableCss | C# | .NET | inline CSS | HTML export | Excel to HTML | no external CSS | self‑contained HTML | Windows
// Common Searches: Aspose.Cells export Excel to HTML without CSS file | HtmlSaveOptions.DisableCss C# example | how to embed styles inline when saving workbook as HTML | C# generate HTML from Excel with inline styles only | disable external CSS Aspose.Cells HTML export
// Developer Intent: Export an Excel workbook to HTML using Aspose.Cells while embedding all styling inline and preventing the creation of a separate CSS file.
// Use Cases: Send a single‑file HTML report via email without attaching a stylesheet. | Embed spreadsheet snapshots in web pages where external resources are blocked. | Create printable HTML output that retains formatting without extra files.
// AI Prompts: Show a C# code snippet that saves an Aspose.Cells workbook as HTML with inline CSS only using HtmlSaveOptions.DisableCss. | Explain why HtmlSaveOptions.DisableCss is needed and how it affects the generated HTML. | Provide step‑by‑step instructions to export Excel to a self‑contained HTML page with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // A C# example that creates a workbook, applies cell formatting, sets HtmlSaveOptions.DisableCss = true, and saves the file as a self‑contained HTML page with all styles embedded inline—no external CSS file is generated.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with some formatting
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            Style style = sheet.Cells["A1"].GetStyle();
            style.Font.IsBold = true;
            style.Font.Color = System.Drawing.Color.Blue;
            sheet.Cells["A1"].SetStyle(style);

            // Configure HTML save options to use only inline styles (disable external CSS)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableCss = true; // Inline styles only, no external CSS file

            // Define output path
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "HtmlWithInlineStyles.html");

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
