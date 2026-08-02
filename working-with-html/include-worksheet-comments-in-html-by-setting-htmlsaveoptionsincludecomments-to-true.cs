// Title: How to Include Worksheet Comments in HTML Output Using Aspose.Cells for .NET
// Description: Shows how to add a comment to a cell, enable comment export with HtmlSaveOptions (IsExportComments = true), and save the workbook as an HTML file where the comments are rendered alongside the data.
// Keywords: Aspose.Cells | C# HTML export | worksheet comments | HtmlSaveOptions | IsExportComments | export comments to HTML | Excel to HTML conversion | Aspose.Cells example | save workbook as HTML | include cell notes
// Common Searches: Aspose.Cells include comments when saving to HTML | HtmlSaveOptions IsExportComments property C# | export Excel cell notes to HTML with Aspose | C# convert workbook to HTML with comments | how to show worksheet comments in HTML using Aspose.Cells
// Developer Intent: Create an HTML representation of an Excel workbook that preserves all cell comments.
// Use Cases: Publish Excel reports on a website while keeping annotation information visible. | Provide a read‑only web view of a spreadsheet that includes reviewer notes. | Automate batch conversion of workbooks to HTML for documentation portals, ensuring comments are not lost.
// AI Prompts: Generate C# code that loads an existing workbook, adds comments to several cells, and saves it as HTML with comments using Aspose.Cells. | Explain the effect of the IsExportComments flag on the generated HTML and describe which HTML elements contain the comment text. | Provide a step‑by‑step guide to customize the appearance of exported comments (font, color, tooltip) in the HTML output.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentExport
{
    // Shows how to add a comment to a cell, enable comment export with HtmlSaveOptions (IsExportComments = true), and save the workbook as an HTML file where the comments are rendered alongside the data.
    public class ExportCommentsToHtml
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to a cell
            worksheet.Cells["A1"].PutValue("Sample Data");

            // Add a comment to the same cell
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This is a test comment";

            // Configure HTML save options to include comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true // Enable exporting of worksheet comments
            };

            // Save the workbook as HTML with comments included
            workbook.Save("output_with_comments.html", htmlOptions);

            Console.WriteLine("HTML file with comments exported successfully.");
        }
    }
}
