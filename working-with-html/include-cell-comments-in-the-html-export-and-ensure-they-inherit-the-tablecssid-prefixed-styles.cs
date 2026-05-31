using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExportWithComments
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);

            // Add a comment to cell A2
            int commentIndex = sheet.Comments.Add("A2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Fresh and organic";

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export comments so they appear in the HTML output
                IsExportComments = true,
                // Use in‑place printing of comments (appears next to the cell)
                ExportCommentsType = PrintCommentsType.PrintInPlace,
                // Prefix for CSS class names generated for table elements
                TableCssId = "custom-table",
                // Optional: add a style that will be applied to comment elements
                // (comments will inherit the table CSS prefix because they are rendered
                // inside the same table structure)
                CssStyles = @"
                    /* Example style for comments */
                    .custom-table .comment
                    {
                        background-color:#ffffe0;
                        border:1px solid #c0c0c0;
                        padding:4px;
                        font-style:italic;
                    }"
            };

            // Save the workbook as an HTML file with the configured options
            workbook.Save("WorkbookWithComments.html", htmlOptions);

            Console.WriteLine("HTML file with comments and TableCssId styles has been created.");
        }
    }
}