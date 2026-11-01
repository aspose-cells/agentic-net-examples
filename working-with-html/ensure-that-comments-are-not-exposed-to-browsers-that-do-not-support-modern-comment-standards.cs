// Title: Hide Excel cell comments when exporting a workbook to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates an Excel workbook, adds a hidden comment to a cell, and saves it as HTML using Aspose.Cells without including the comment in the output. | Show how to configure Aspose.Cells HtmlSaveOptions and set Comment.IsVisible to false so that comments are omitted from the generated HTML file. | Write a complete example that creates a worksheet, inserts data, adds a comment, hides it, ensures the output folder exists, and exports the workbook to HTML.
// Common Searches: Aspose.Cells C# export to HTML without showing cell comments | How to prevent Excel comments from appearing in HTML output using Aspose.Cells | C# hide comment visibility when saving workbook as HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions hide comments example | Export Excel file to HTML and exclude comments in .NET
// Tags: Aspose.Cells hide comment HTML export | HtmlSaveOptions comment visibility C# | export workbook to HTML without comments | Excel comment.IsVisible false Aspose.Cells | C# create workbook and hide cell comment

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, adds data to cells A1 and B1, inserts a comment on A1, sets the comment's IsVisible property to false, configures HtmlSaveOptions for HTML format, ensures the output directory exists, and saves the workbook as an HTML file while suppressing the comment from the generated markup.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a comment that should not appear in the HTML output.";
            // Hide the comment so it won't be exported
            comment.IsVisible = false;

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Determine output path and ensure directory exists
            string outputPath = "output.html";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
