using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello World");

        // Add a comment to demonstrate comment export (optional)
        Comment comment = sheet.Comments[sheet.Comments.Add("A1")];
        comment.Note = "Sample comment";

        // Configure HTML save options to disable downlevel‑revealed conditional comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            DisableDownlevelRevealedComments = true
        };

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}

// Author: Example showing how to disable downlevel‑revealed comments for older browsers.