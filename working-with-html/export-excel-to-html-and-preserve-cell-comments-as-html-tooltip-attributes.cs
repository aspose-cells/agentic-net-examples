// Title: Export Excel to HTML with cell comments as tooltip attributes using Aspose.Cells for .NET
// Description: Shows how to build a workbook, add comments, set HtmlSaveOptions (IsExportComments = true, ExportCommentsType = PrintInPlace) and save it as HTML so each comment appears as a title‑attribute tooltip on the corresponding cell.
// Keywords: Aspose.Cells | C# | HTML export | cell comments | tooltip | IsExportComments | PrintInPlace | HtmlSaveOptions | Excel to HTML conversion | preserve comments
// Common Searches: Aspose.Cells export comments as HTML tooltip | How to keep Excel comments when converting to HTML in C# | PrintCommentsType.PrintInPlace example | Save workbook as HTML with tooltips | C# Aspose.Cells HtmlSaveOptions comment settings
// Developer Intent: Convert an Excel workbook to HTML while retaining cell comments as hover tooltips.
// Use Cases: Publish a product catalog where reviewer notes become on‑hover tooltips for each item. | Generate financial dashboards that show explanatory comments as HTML title attributes without extra UI components. | Automate documentation pipelines that transform internal Excel sheets into web pages while preserving reviewer feedback.
// AI Prompts: Modify the example to output comments as footnotes instead of tooltips in the generated HTML. | Provide code that writes the HTML to a MemoryStream and returns the string for further processing. | Explain how to attach a custom CSS class to the tooltip elements created from cell comments.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtmlWithComments
{
    // Shows how to build a workbook, add comments, set HtmlSaveOptions (IsExportComments = true, ExportCommentsType = PrintInPlace) and save it as HTML so each comment appears as a title‑attribute tooltip on the corresponding cell.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // Add comments to some cells
            Comment commentA2 = sheet.Comments[sheet.Comments.Add("A2")];
            commentA2.Note = "Fresh apples from the orchard";

            Comment commentB3 = sheet.Comments[sheet.Comments.Add("B3")];
            commentB3.Note = "Discounted price for bananas";

            // Configure HTML save options to export comments as tooltips
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Enable exporting of comments
                IsExportComments = true,
                // Print comments in place so they appear as HTML title attributes (tooltips)
                ExportCommentsType = PrintCommentsType.PrintInPlace
            };

            // Save the workbook as HTML with the configured options
            workbook.Save("Products.html", htmlOptions);

            Console.WriteLine("HTML file 'Products.html' created with comments preserved as tooltips.");
        }
    }
}
