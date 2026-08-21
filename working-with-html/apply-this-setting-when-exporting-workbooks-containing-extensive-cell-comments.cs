// Title: Export Workbooks with Thousands of Cell Comments to HTML using Aspose.Cells for .NET
// Description: The sample builds a workbook, inserts 1,000 comments across 100 rows × 10 columns, enables HtmlSaveOptions.IsExportComments, optionally sets ExportCommentsType, and saves the result as an HTML file that retains every comment.
// Keywords: Aspose.Cells HTML comment export | IsExportComments | ExportCommentsType | save workbook as HTML with comments | bulk comment export .NET | C# Aspose.Cells export Excel comments | HTMLSaveOptions | large comment dataset Aspose.Cells
// Common Searches: Aspose.Cells export comments to HTML | Include cell comments when saving as HTML in .NET | HTMLSaveOptions IsExportComments example | Export thousands of Excel comments to HTML | Set ExportCommentsType in Aspose.Cells | C# save workbook with comments as HTML
// Developer Intent: Generate an HTML version of an Excel file that contains all cell comments using Aspose.Cells for .NET.
// Use Cases: Create a web‑ready report that shows every annotation from a heavily commented worksheet. | Archive an Excel document with extensive reviewer notes in a format viewable in browsers. | Provide stakeholders with an HTML preview that preserves comment context for quality‑control processes.
// AI Prompts: Show how to export only comments from a specific author to HTML with Aspose.Cells. | Give C# code that styles exported comments using custom CSS in the HTML output. | Explain the effect of setting ExportCommentsType to PrintInPlace versus PrintInSeparateFile.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentExportDemo
{
    // The sample builds a workbook, inserts 1,000 comments across 100 rows × 10 columns, enables HtmlSaveOptions.IsExportComments, optionally sets ExportCommentsType, and saves the result as an HTML file that retains every comment.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a large number of comments to simulate extensive comments
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    // Add a comment to the current cell
                    int commentIndex = sheet.Comments.Add(row, col);
                    Comment comment = sheet.Comments[commentIndex];
                    comment.Note = $"Comment at {CellsHelper.CellIndexToName(row, col)}";
                    comment.Author = "DemoUser";
                }
            }

            // Configure HTML save options to export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Enable exporting of comments (default is false)
                IsExportComments = true,

                // Optionally, specify how comments are exported
                // ExportCommentsType = PrintCommentsType.PrintInPlace
            };

            // Save the workbook as HTML with comments exported
            workbook.Save("WorkbookWithComments.html", htmlOptions);
        }
    }
}
