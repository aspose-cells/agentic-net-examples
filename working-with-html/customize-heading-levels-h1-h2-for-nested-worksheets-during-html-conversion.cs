// Title: Set HTML Heading Levels (h1, h2) for Nested Worksheets with Aspose.Cells for .NET
// Description: Demonstrates how to export a multi‑sheet workbook to HTML where the first worksheet is rendered as an <h1> tag and subsequent worksheets as <h2> tags. Uses HtmlSaveOptions.ExportRowColumnHeadings and custom CSS to style the headings.
// Keywords: Aspose.Cells HTML export | ExportRowColumnHeadings | custom CSS h1 h2 | nested worksheets HTML | C# Aspose.Cells heading tags | global | United States
// Common Searches: Aspose.Cells export sheet names as h1 h2 | how to apply CSS to worksheet headings in HTML output | set different heading levels for multiple sheets Aspose.Cells | property to control heading tags when saving to HTML
// Developer Intent: Generate HTML from an Excel workbook where each worksheet appears with a specific heading level and custom styling.
// Use Cases: Create a summary‑detail report where the summary sheet is an <h1> and detail sheets are <h2> elements. | Produce a financial statement HTML file with distinct heading styles for overview and subsidiary sections. | Convert Excel documentation into web pages with hierarchical headings based on worksheet order.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as HTML with the first sheet as <h1> and all other sheets as <h2>, including CSS for font and color. | Explain how ExportRowColumnHeadings and CssStyles in HtmlSaveOptions work together to define heading tags and styling during HTML conversion. | Show how to assign h1, h2, h3 tags to worksheets programmatically based on their position when exporting to HTML with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlHeadingDemo
{
    // Demonstrates how to export a multi‑sheet workbook to HTML where the first worksheet is rendered as an <h1> tag and subsequent worksheets as <h2> tags. Uses HtmlSaveOptions.ExportRowColumnHeadings and custom CSS to style the headings.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add data to the first worksheet (will be rendered as <h1>)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Summary";
            sheet1.Cells["A1"].PutValue("Report Summary");
            sheet1.Cells["A2"].PutValue("Total Sales");
            sheet1.Cells["B2"].PutValue(12500);

            // Add a second worksheet (will be rendered as <h2>)
            Worksheet sheet2 = workbook.Worksheets.Add("Details");
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Quantity");
            sheet2.Cells["A2"].PutValue("Apples");
            sheet2.Cells["B2"].PutValue(150);
            sheet2.Cells["A3"].PutValue("Oranges");
            sheet2.Cells["B3"].PutValue(200);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export sheet names as headings
                ExportRowColumnHeadings = true,

                // Custom CSS to style the generated <h1> and <h2> tags
                CssStyles = @"
                    h1 {font-family:Arial; font-size:28px; color:#2E4053; margin-bottom:10px;}
                    h2 {font-family:Arial; font-size:22px; color:#566573; margin-top:20px; margin-bottom:8px;}
                "
            };

            // Save the workbook as HTML
            workbook.Save("NestedSheets.html", saveOptions);
        }
    }
}
