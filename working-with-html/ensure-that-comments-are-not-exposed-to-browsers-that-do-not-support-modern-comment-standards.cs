// Title: Export Excel workbook to HTML without cell comments using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML without displaying cell comments | How to hide Excel comments when saving as HTML using Aspose.Cells | Generate HTML from Excel with Aspose.Cells while omitting comments | C# Aspose.Cells HtmlSaveOptions prevent comment export | Exclude hidden comments from HTML output in Aspose.Cells .NET
// Tags: Aspose.Cells HTML export hide comments | C# set comment IsVisible false Aspose | HtmlSaveOptions omit cell comments | Export Excel to HTML without comment data | Workbook to HTML without comments Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, adds a comment to cell A1, sets the comment's IsVisible property to false, configures HtmlSaveOptions, and saves the file as HTML, ensuring that the comment is omitted from the generated HTML output.
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
            // Hide the comment so it won't be exported to HTML
            comment.IsVisible = false;

            // Configure HTML save options (no ExportComments property in this version)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
