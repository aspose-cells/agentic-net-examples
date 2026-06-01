using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCommentExportValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Sample Data");

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a test comment";

            // Configure HTML save options to export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true   // Enable comment export
            };

            // Define output HTML file path
            string htmlPath = "output_with_comments.html";

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlPath, htmlOptions);

            // Read the generated HTML file
            string htmlContent = File.ReadAllText(htmlPath);

            // Validate that the HTML contains comment tags (<!--)
            bool containsCommentTag = htmlContent.Contains("<!--");

            // Output the validation result
            Console.WriteLine($"HTML contains comment tag: {containsCommentTag}");
        }
    }
}