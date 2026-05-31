using System;
using System.IO;
using Aspose.Cells;

class ExportCommentDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to cell A1
        worksheet.Cells["A1"].PutValue("Hello World");

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "This is a test comment";

        // Set HTML save options to export comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            IsExportComments = true
        };

        // Save the workbook as HTML
        string htmlFile = "output_with_comments.html";
        workbook.Save(htmlFile, htmlOptions);

        // Verify that the comment appears in the generated HTML
        string htmlContent = File.ReadAllText(htmlFile);
        bool commentFound = htmlContent.Contains("This is a test comment");
        Console.WriteLine("Comment exported to HTML: " + commentFound);
    }
}