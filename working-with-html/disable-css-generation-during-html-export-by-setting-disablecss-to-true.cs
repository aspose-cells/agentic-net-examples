// Title: Disable CSS Generation (Use Inline Styles) When Exporting Excel to HTML with Aspose.Cells for .NET
// Description: This C# example creates an in‑memory workbook, applies bold blue formatting to cell A1, sets HtmlSaveOptions.DisableCss to true, and saves the workbook as a self‑contained HTML file that uses only inline style attributes, eliminating external CSS files.
// Keywords: Aspose.Cells HtmlSaveOptions DisableCss | export Excel to HTML inline styles | disable CSS generation Aspose.Cells .NET | HTML export without external stylesheet | C# Aspose.Cells HTML inline styling
// Common Searches: Aspose.Cells disable CSS when exporting to HTML | HtmlSaveOptions.DisableCss example C# | export Excel workbook to HTML with inline styles | remove external CSS from Aspose.Cells HTML output | self‑contained HTML from spreadsheet Aspose.Cells
// Developer Intent: Generate HTML from a workbook while suppressing external CSS and using only inline style attributes.
// Use Cases: Create email‑ready HTML reports from spreadsheets without linking to CSS files. | Produce lightweight, single‑file HTML previews for web applications that restrict external resources. | Embed styled spreadsheet data in documentation or blogs where only inline CSS is allowed.
// AI Prompts: Show how to set HtmlSaveOptions.DisableCss to true in Aspose.Cells for .NET. | Provide C# code that exports an Excel workbook to HTML using only inline styles. | Explain the impact of disabling CSS generation on the HTML output produced by Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // This C# example creates an in‑memory workbook, applies bold blue formatting to cell A1, sets HtmlSaveOptions.DisableCss to true, and saves the workbook as a self‑contained HTML file that uses only inline style attributes, eliminating external CSS files.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Add some sample data to demonstrate styling
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A1"].GetStyle().Font.IsBold = true;
            sheet.Cells["A1"].GetStyle().Font.Color = System.Drawing.Color.Blue;

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable CSS generation – use only inline styles
            htmlOptions.DisableCss = true;

            // Save the workbook as HTML with the specified options
            workbook.Save("HtmlWithInlineStyles.html", htmlOptions);

            Console.WriteLine("Workbook saved as HTML with inline styles only.");
        }
    }
}
