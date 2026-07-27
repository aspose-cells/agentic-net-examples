// Title: Export Workbooks with Hundreds of Cell Comments to HTML using Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds 100 visible comments to cells A2‑A101, configures HtmlSaveOptions (IsExportComments, ExportCommentsType=PrintInPlace, ExcludeUnusedStyles, PageTitle) and saves the file as HTML. Demonstrates handling large comment collections and optimizing output performance.
// Keywords: Aspose.Cells | C# | .NET | HTML export | cell comments | IsExportComments | PrintInPlace | ExcludeUnusedStyles | custom page title | large comment set | performance optimization
// Common Searches: Aspose.Cells export comments to HTML | How to export many cell comments with Aspose.Cells | HtmlSaveOptions IsExportComments example C# | PrintCommentsType PrintInPlace usage | ExcludeUnusedStyles large workbook performance | Set HTML page title Aspose.Cells | C# export spreadsheet comments to HTML
// Developer Intent: Generate an HTML file from a .NET workbook that contains a large number of visible cell comments, preserving their placement and minimizing file size.
// Use Cases: Create web‑ready reports that display every spreadsheet comment next to its cell. | Improve load time for HTML reports generated from workbooks with hundreds of comments. | Add SEO‑friendly page titles when publishing spreadsheet data as HTML. | Integrate into automated reporting pipelines where comment visibility is required.
// AI Prompts: Write C# code using Aspose.Cells to add 100 comments and export the workbook to HTML with comments rendered in‑place. | Explain which HtmlSaveOptions settings speed up HTML export for workbooks that contain many comments. | Show how to set a custom HTML page title and exclude unused CSS styles during the export process. | Provide a step‑by‑step guide for exporting a spreadsheet with extensive comments while keeping them visible in the browser.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentExportDemo
{
    // C# example that creates a workbook, adds 100 visible comments to cells A2‑A101, configures HtmlSaveOptions (IsExportComments, ExportCommentsType=PrintInPlace, ExcludeUnusedStyles, PageTitle) and saves the file as HTML. Demonstrates handling large comment collections and optimizing output performance.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Data with many comments");

            // Add a large number of comments to simulate an extensive comment scenario
            // Here we add comments to cells A2 to A101 (100 comments)
            for (int i = 2; i <= 101; i++)
            {
                // Add a comment to cell A{i}
                int commentIndex = sheet.Comments.Add(0, i - 1); // column 0 (A), row i-1 (zero‑based)
                Comment comment = sheet.Comments[commentIndex];
                comment.Note = $"Comment number {i - 1}";
                comment.Author = "DemoUser";
                comment.IsVisible = true;
            }

            // Configure HTML save options to export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Enable exporting of comments
                IsExportComments = true,
                // Choose how comments are rendered in the HTML (in‑place display)
                ExportCommentsType = PrintCommentsType.PrintInPlace,
                // Optional: improve performance for large files by excluding unused styles
                ExcludeUnusedStyles = true,
                // Optional: set a page title for the generated HTML
                PageTitle = "Workbook with Extensive Comments"
            };

            // Save the workbook to an HTML file using the configured options
            workbook.Save("WorkbookWithComments.html", htmlOptions);
        }
    }
}
