// Title: Export Cell Comments to HTML using Aspose.Cells for .NET
// Description: Shows how to add a comment to a worksheet cell, turn on the IsExportComments flag in HtmlSaveOptions, and save the workbook as an HTML file so the comment is rendered in the web page.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | IsExportComments | export comments to HTML | cell comment HTML | Excel to HTML conversion | Aspose.Cells example | GitHub snippet
// Common Searches: Aspose.Cells export cell comments to HTML | How to enable IsExportComments in HtmlSaveOptions | C# save Excel as HTML with comments | Include Excel comments in HTML output Aspose | Aspose.Cells HTML conversion example | GitHub Aspose.Cells comment export sample
// Developer Intent: Include worksheet cell comments in the generated HTML file.
// Use Cases: Create web‑ready reports that retain Excel note annotations. | Build documentation portals where reviewers see original comments alongside data. | Automate batch conversion of spreadsheets to HTML while preserving comment metadata for compliance audits. | Integrate comment‑aware HTML export into ASP.NET web applications.
// AI Prompts: Generate a C# code snippet that saves a workbook to HTML with comments using Aspose.Cells. | Explain the effect of HtmlSaveOptions.IsExportComments on the resulting HTML file. | Show how to customize comment appearance (font, color) when exporting to HTML with Aspose.Cells. | Provide steps to retrieve comment author and text after loading an HTML file back into Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to add a comment to a worksheet cell, turn on the IsExportComments flag in HtmlSaveOptions, and save the workbook as an HTML file so the comment is rendered in the web page.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some data to a cell
        sheet.Cells["A1"].PutValue("Sample Data");

        // Add a comment to the same cell
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is a test comment";
        comment.Author = "Aspose";

        // Configure HTML save options to export comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            IsExportComments = true   // Enable exporting of comments
        };

        // Save the workbook as HTML with comments included
        workbook.Save("output_with_comments.html", htmlOptions);
    }
}
