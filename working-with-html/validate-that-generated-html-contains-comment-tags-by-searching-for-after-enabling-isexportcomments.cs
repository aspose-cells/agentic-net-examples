// Title: Export a worksheet cell comment to HTML and verify the presence of <!-- comment tags using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a comment to cell A1, saves the workbook as HTML with HtmlSaveOptions.IsExportComments enabled, and returns a boolean indicating if the generated HTML contains the '<!--' marker. | Show how to read the HTML string from a MemoryStream after saving with Aspose.Cells and programmatically check for HTML comment tags.
// Common Searches: Aspose.Cells C# export cell comments to HTML and detect <!-- tag | how to enable comment export in Aspose.Cells HtmlSaveOptions | validate that HTML output from Aspose.Cells contains comment markers | C# read HTML from MemoryStream after Aspose.Cells save and check for comments | saving Excel workbook as HTML with comments using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions export comments | C# export Excel comments to HTML | validate HTML comment tags Aspose.Cells | memory stream HTML output Aspose.Cells | IsExportComments property usage

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds a comment to cell A1, saves the workbook as HTML with IsExportComments set to true, reads the HTML from a MemoryStream, and determines whether the output includes an HTML comment tag (<!--).
class ExportCommentsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add a comment to cell A1
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World");

            // Add comment and set its note
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a comment";

            // Configure HTML save options to export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true // Enable comment export
            };

            // Save the workbook to a memory stream as HTML
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0; // Reset stream position for reading

                // Read the generated HTML into a string
                string htmlContent;
                using (StreamReader reader = new StreamReader(htmlStream))
                {
                    htmlContent = reader.ReadToEnd();
                }

                // Validate that the HTML contains comment tags (<!--)
                bool containsCommentTag = htmlContent.Contains("<!--");
                Console.WriteLine("HTML contains comment tags: " + containsCommentTag);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
