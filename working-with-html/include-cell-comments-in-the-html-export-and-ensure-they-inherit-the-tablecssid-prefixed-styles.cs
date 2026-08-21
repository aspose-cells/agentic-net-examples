// Title: Export Cell Comments to HTML with TableCssId‑Prefixed Styles Using Aspose.Cells for .NET (C#)
// Description: Shows how to add a comment to a worksheet cell, enable comment export, assign a custom TableCssId, and save the workbook as HTML so that comments are rendered and inherit the prefixed CSS styling.
// Keywords: Aspose.Cells | C# | .NET | HTML export | cell comments | TableCssId | custom CSS id | HtmlSaveOptions | export comments | Excel to HTML
// Common Searches: Aspose.Cells export comments to HTML | How to use TableCssId in HtmlSaveOptions | Include cell notes in HTML output Aspose | C# export worksheet with comments | Custom CSS for Aspose.Cells HTML
// Developer Intent: Create an HTML representation of a workbook that includes cell comments and applies a user‑defined CSS identifier to the generated table.
// Use Cases: Building web‑based reports that show Excel comments alongside data. | Embedding spreadsheet content in documentation with consistent styling via a custom TableCssId. | Generating interactive HTML pages where comments follow the same CSS rules as the table layout.
// AI Prompts: Modify the example so that exported comments appear as inline text instead of tooltips. | Provide CSS snippets that target the TableCssId‑prefixed selectors to style comment bubbles. | Show how to export comments for only selected cells while keeping the TableCssId styling intact.

using System;
using Aspose.Cells;

// Shows how to add a comment to a worksheet cell, enable comment export, assign a custom TableCssId, and save the workbook as HTML so that comments are rendered and inherit the prefixed CSS styling.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data
        worksheet.Cells["A1"].PutValue("Hello World");

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "This is a sample comment that will appear in the HTML export.";

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Enable exporting of comments
            IsExportComments = true,
            // Alternatively you can use:
            // ExportCommentsType = PrintCommentsType.PrintInPlace,

            // Set a prefix for CSS classes inside the generated table
            TableCssId = "custom-table"
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", saveOptions);
    }
}
