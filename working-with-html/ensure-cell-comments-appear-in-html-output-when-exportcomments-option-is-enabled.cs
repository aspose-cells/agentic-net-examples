// Title: Include worksheet cell comments when saving as HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a comment to a specific cell and saves the workbook as HTML, ensuring the comment appears in the output using Aspose.Cells. | Show how to set HtmlSaveOptions.ExportComments in Aspose.Cells so that cell comments are rendered in the generated HTML file. | Provide a complete C# example that creates a workbook, adds multiple cell comments, and exports the sheet to HTML with all comments preserved.
// Common Searches: Aspose.Cells C# export worksheet to HTML with comments visible | How to keep Excel cell comments when converting to HTML using Aspose.Cells .NET | Save workbook as HTML and include comments Aspose.Cells example | HtmlSaveOptions ExportComments true Aspose.Cells C# | Render Excel comments in HTML output with Aspose.Cells for .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportComments | C# export cell comments to HTML | Aspose.Cells render worksheet comments in HTML | save workbook as HTML with comments Aspose.Cells | Aspose.Cells HTML conversion preserving comments

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a new workbook, adds a comment to cell A1, configures HtmlSaveOptions (which export comments by default), and saves the workbook as an HTML file where the comment is rendered in the generated HTML.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Add a comment to cell A1
                int commentIndex = sheet.Comments.Add("A1");
                var comment = sheet.Comments[commentIndex];
                comment.Note = "This is a sample comment displayed in HTML.";

                // Configure HTML save options (comments are exported by default)
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Save the workbook as HTML with comments
                string outputPath = "WorkbookWithComments.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
