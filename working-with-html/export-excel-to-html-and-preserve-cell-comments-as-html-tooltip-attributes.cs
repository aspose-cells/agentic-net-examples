// Title: Export Excel to HTML with cell comments as tooltips using Aspose.Cells for .NET
// Description: Shows how to generate an HTML file from an Excel workbook where each cell comment is rendered as a hover‑over tooltip. The example creates a workbook, adds data and a comment, then saves with HtmlSaveOptions (IsExportComments = true, ExportCommentsType = PrintInPlace).
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | HTML export | cell comments | tooltip | IsExportComments | PrintInPlace | HtmlSaveOptions | preserve comments | export comments as tooltip
// Common Searches: Aspose.Cells export Excel comments to HTML tooltip | How to keep cell comments when saving as HTML in .NET | HtmlSaveOptions IsExportComments example | Render Excel comments as HTML title attribute | C# export workbook to HTML with comments
// Developer Intent: Create an HTML representation of an Excel sheet that displays cell comments as hover‑over tooltips.
// Use Cases: Publish a price list online where discount notes appear as tooltips on price cells. | Convert internal review spreadsheets to web pages while retaining reviewer comments for accessibility. | Automate dashboard publishing from Excel to web format without losing comment information.
// AI Prompts: Modify the code to output comments as footnotes instead of tooltips. | Provide a sample that loads an existing .xlsx, adds several comments, and saves to HTML with tooltip comments. | Explain how to customize the generated HTML to change the tooltip style or use a custom JavaScript library.

using System;
using Aspose.Cells;

namespace ExportExcelToHtmlWithComments
{
    // Shows how to generate an HTML file from an Excel workbook where each cell comment is rendered as a hover‑over tooltip. The example creates a workbook, adds data and a comment, then saves with HtmlSaveOptions (IsExportComments = true, ExportCommentsType = PrintInPlace).
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to cells
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // Add a comment to a cell
            int commentIndex = sheet.Comments.Add("B2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Discounted price for today";

            // Configure HTML save options to export comments as tooltips
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true,                     // Enable comment export
                ExportCommentsType = PrintCommentsType.PrintInPlace // Render comments as tooltips
            };

            // Save the workbook as an HTML file with comments preserved
            string outputPath = "WorkbookWithComments.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
