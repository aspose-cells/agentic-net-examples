// Title: Export Cell Comments to HTML with Aspose.Cells for .NET (IsExportComments)
// Description: Shows how to insert a comment into a worksheet cell, enable HtmlSaveOptions.IsExportComments, and save the workbook as an HTML file that displays the comment as a tooltip or visible note.
// Keywords: Aspose.Cells | C# export comments HTML | HtmlSaveOptions IsExportComments | cell comment to HTML | Aspose.Cells HTML output | export Excel comments | .NET workbook to HTML | HTML report with comments | save workbook as HTML | comment tooltip rendering
// Common Searches: Aspose.Cells include cell comments in HTML export | HtmlSaveOptions IsExportComments C# example | How to show Excel comments in HTML using Aspose | Export Excel worksheet with comments to web page | Enable comment tooltips in Aspose.Cells HTML output
// Developer Intent: Render cell comments in the HTML file produced by Aspose.Cells.
// Use Cases: Create an interactive HTML report where reviewers can see comments as hover‑over notes. | Publish Excel data on an intranet with comments preserved for context. | Automate documentation pipelines that convert spreadsheets with annotations into web‑ready pages.
// AI Prompts: Generate a C# example that saves a workbook to HTML with comments styled using custom CSS in Aspose.Cells. | Explain how to toggle comment visibility modes (tooltip vs. inline) via HtmlSaveOptions in Aspose.Cells. | Show code to export each worksheet’s comments to separate HTML files while keeping the main data file.

using System;
using Aspose.Cells;

// Shows how to insert a comment into a worksheet cell, enable HtmlSaveOptions.IsExportComments, and save the workbook as an HTML file that displays the comment as a tooltip or visible note.
class ExportCommentsToHtml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some data to a cell
        sheet.Cells["A1"].PutValue("Sample Data");

        // Add a comment to the same cell
        int commentIdx = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIdx];
        comment.Author = "Demo Author";
        comment.Note = "This is a test comment.";

        // Set HTML save options to export comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            IsExportComments = true // Enable comment export
        };

        // Save the workbook as HTML with comments included
        workbook.Save("output_with_comments.html", htmlOptions);
    }
}
