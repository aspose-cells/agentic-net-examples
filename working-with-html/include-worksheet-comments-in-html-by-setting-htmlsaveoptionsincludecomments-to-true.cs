// Title: Aspose.Cells for .NET – Export Worksheet Comments to HTML with HtmlSaveOptions.IncludeComments
// Description: Demonstrates how to add a comment to a cell, enable HtmlSaveOptions.IncludeComments, and save the workbook as an HTML file that shows the comment as a tooltip or inline note.
// Keywords: Aspose.Cells export comments HTML | HtmlSaveOptions IncludeComments .NET | C# save Excel as HTML with comments | worksheet comment to HTML Aspose | Excel comments in web view
// Common Searches: include cell comments when saving Excel to HTML Aspose.Cells | HtmlSaveOptions IncludeComments example C# | export Excel comments to HTML using Aspose | Aspose.Cells HTML output with notes | how to show worksheet comments in HTML
// Developer Intent: Generate an HTML representation of an Excel workbook that retains the original cell comments.
// Use Cases: Create web‑ready reports that display analyst notes attached to cells. | Publish interactive spreadsheets where comments appear as hover tooltips. | Automate documentation pipelines that require both data and its annotations in HTML format.
// AI Prompts: Show how to style exported comments (font, color, background) in the HTML output with Aspose.Cells. | Provide code to export multiple worksheets, each preserving its comments, into a single HTML file. | Explain how to convert comment positions to absolute coordinates for custom JavaScript tooltip handling.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentsToHtml
{
    // Demonstrates how to add a comment to a cell, enable HtmlSaveOptions.IncludeComments, and save the workbook as an HTML file that shows the comment as a tooltip or inline note.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Hello World");

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment";

            // Configure HTML save options to include comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true   // Enable exporting of worksheet comments
            };

            // Save the workbook as HTML with comments included
            string outputPath = "WorkbookWithComments.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
