// Title: Export an Excel worksheet to HTML with gridlines and cell comments using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to HTML with gridlines visible | How to include cell comments when saving Excel as HTML using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportGridLines true and preserve comments example | Generate HTML from Excel file with gridlines and comments in .NET | Save first worksheet as HTML with comments and gridlines using Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines property | Excel worksheet HTML export with cell comments | C# HTML conversion preserving Excel gridlines | Aspose.Cells HTML output including comments | save Excel as HTML with gridlines C#

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an Excel workbook, makes gridlines visible, ensures a comment exists in cell A1, configures HtmlSaveOptions with ExportGridLines enabled (comments are exported by default), and saves the first worksheet as an HTML file preserving both gridlines and cell comments.
class ExportWorksheetToHtml
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure gridlines are visible in the worksheet
            sheet.IsGridlinesVisible = true;

            // Add a sample comment (or reuse the first existing one)
            Comment comment;
            if (sheet.Comments.Count > 0)
            {
                comment = sheet.Comments[0];
            }
            else
            {
                // Add returns the index of the new comment; retrieve it afterwards
                int commentIndex = sheet.Comments.Add("A1");
                comment = sheet.Comments[commentIndex];
            }
            comment.Note = "This is a sample comment.";

            // Configure HTML save options to export gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true // Include gridlines in the HTML output
                // Comments are exported by default; no explicit property needed
            };

            // Save the first worksheet as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
