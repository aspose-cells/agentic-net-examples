// Title: Export cell comments to HTML using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export cell comments when saving as HTML | How to include Excel comments in HTML output using Aspose.Cells | Enable comment export in HtmlSaveOptions for Aspose.Cells .NET | Save workbook to HTML with comments displayed using Aspose.Cells | Aspose.Cells HTML export preserving cell notes C#
// Tags: Aspose.Cells HtmlSaveOptions comment export | C# export Excel cell notes to HTML | include worksheet comments in HTML output Aspose | save workbook as HTML preserving comments | Aspose.Cells HTML conversion with comments

using System;
using Aspose.Cells;

// The example creates a workbook, adds a comment to cell A1, configures HtmlSaveOptions to export comments, and saves the workbook as an HTML file so the comment appears in the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment.";
            comment.Author = "Author";

            // Configure HTML save options (comments are exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML with comments exported
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
