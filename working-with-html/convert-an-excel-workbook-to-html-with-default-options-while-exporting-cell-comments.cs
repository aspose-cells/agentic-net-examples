// Title: Convert Excel Workbook to HTML with Cell Comments using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a comment to a cell, enable comment export via HtmlSaveOptions, and save the file as HTML with default settings, preserving the comments in the output.
// Keywords: Aspose.Cells | C# HTML conversion | export Excel comments | HtmlSaveOptions IsExportComments | Excel to HTML .NET | cell comment rendering | default save options | workbook.Save HTML | Aspose.Cells tutorial
// Common Searches: Aspose.Cells export comments to HTML C# | HtmlSaveOptions IsExportComments example | Convert Excel file to HTML with comments | Save workbook as HTML preserving comments | C# code for Excel to HTML conversion Aspose
// Developer Intent: Generate an HTML version of an Excel sheet that includes all cell comments.
// Use Cases: Publish spreadsheet data on a website while keeping reviewer notes visible as tooltips. | Create printable HTML reports that retain the original worksheet comments for audit trails. | Automate batch conversion of multiple Excel files to web‑ready HTML with comment preservation.
// AI Prompts: Write C# code with Aspose.Cells to convert an existing workbook to HTML and include every cell comment. | Show how to modify HtmlSaveOptions to customize the appearance of exported comments in the HTML output. | Explain a method for processing a folder of Excel files into HTML files that keep all comments using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add a comment to a cell, enable comment export via HtmlSaveOptions, and save the file as HTML with default settings, preserving the comments in the output.
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
        comment.Note = "This is a sample comment";

        // Configure HTML save options to export comments (default options otherwise)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            IsExportComments = true // Enable exporting of cell comments
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
