// Title: Export Excel to HTML with inline styles only using AspNet.Cells HtmlSaveOptions.DisableCss (C#)
// Description: Demonstrates how to create a workbook, apply formatting, and save it as a self‑contained HTML file that uses only inline style attributes by setting HtmlSaveOptions.DisableCss to true.
// Keywords: Aspose.Cells HtmlSaveOptions.DisableCss | C# export Excel to HTML inline styles | disable external CSS Aspose.Cells | self‑contained HTML from Excel | Aspose.Cells HTML export without CSS file
// Common Searches: Aspose.Cells disable external CSS when saving HTML | C# generate HTML with inline styles from Excel | HtmlSaveOptions.DisableCss example | export workbook to HTML without CSS file .NET | inline styling Aspose.Cells HTML export
// Developer Intent: Export an Excel workbook to HTML that contains only inline style attributes, eliminating any separate CSS file.
// Use Cases: Email‑ready HTML reports that need no external stylesheet. | Embedding spreadsheet data in web pages where external CSS loading is prohibited. | Creating portable HTML files for environments with restricted file system access.
// AI Prompts: Provide a C# snippet that saves a workbook as HTML with inline styles only, confirming no .css file is produced. | Explain how HtmlSaveOptions.DisableCss works and show its effect on styled cells. | Show how to combine DisableCss with ExportImagesAsBase64 to generate a fully self‑contained HTML document.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, apply formatting, and save it as a self‑contained HTML file that uses only inline style attributes by setting HtmlSaveOptions.DisableCss to true.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data with formatting
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            // Apply bold style to demonstrate inline styling
            var style = sheet.Cells["A1"].GetStyle();
            style.Font.IsBold = true;
            sheet.Cells["A1"].SetStyle(style);

            // Initialize HTML save options and disable external CSS generation
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableCss = true; // Use only inline styles

            // Save the workbook as an HTML file with the specified options
            workbook.Save("HtmlWithInlineStyles.html", htmlOptions);
        }
    }
}
