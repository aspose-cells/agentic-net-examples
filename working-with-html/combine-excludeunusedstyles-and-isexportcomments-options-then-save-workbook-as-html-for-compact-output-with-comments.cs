// Title: Export a workbook to compact HTML with comments using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a cell comment, and save it as a single lightweight HTML file by enabling ExcludeUnusedStyles and IsExportComments in HtmlSaveOptions.
// Keywords: Aspose.Cells HTML export | ExcludeUnusedStyles option | IsExportComments option | compact HTML workbook | C# Aspose.Cells save as HTML | export cell comments to HTML | minimal CSS Aspose.Cells
// Common Searches: Aspose.Cells save workbook as HTML without unused CSS | include cell comments when exporting to HTML Aspose.Cells | how to generate lightweight HTML from Excel using Aspose.Cells | HtmlSaveOptions ExcludeUnusedStyles and IsExportComments example
// Developer Intent: Generate a minimal‑size HTML file from a spreadsheet that retains cell comments.
// Use Cases: Web preview of spreadsheet data with annotations | Embedding spreadsheet content in documentation or emails | Creating lightweight HTML reports for mobile browsers
// AI Prompts: Show C# code that saves an Aspose.Cells workbook as compact HTML with comments. | Explain the impact of ExcludeUnusedStyles and IsExportComments on HTML output size and content. | Provide step‑by‑step instructions to configure HtmlSaveOptions for minimal CSS and comment export.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, add a cell comment, and save it as a single lightweight HTML file by enabling ExcludeUnusedStyles and IsExportComments in HtmlSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Hello World");

            // Add a comment to the cell
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment";

            // Configure HTML save options:
            // - ExcludeUnusedStyles = true (default) to produce compact HTML
            // - IsExportComments = true to include comments in the output
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExcludeUnusedStyles = true,
                IsExportComments = true
            };

            // Save the workbook as HTML
            string outputPath = "CompactWithComments.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with compact HTML and comments exported.");
        }
    }
}
