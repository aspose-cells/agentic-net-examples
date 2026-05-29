using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Add a comment to demonstrate comment export
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Sample comment";

            // Configure HTML save options:
            // - ExcludeUnusedStyles = true  (compact output, default value)
            // - IsExportComments = true    (include comments in HTML)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExcludeUnusedStyles = true,
                IsExportComments = true
            };

            // Save the workbook as HTML
            string outputPath = "output.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to {outputPath} with comments and unused styles excluded.");
        }
    }
}