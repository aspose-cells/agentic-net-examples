// Title: Export Excel to PDF with Cell Comments as Footnotes using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add cell comments, configure PrintCommentsType.PrintSheetEnd to render those comments as footnotes, optionally set a page footer, and save the worksheet as a PDF with PdfSaveOptions.
// Keywords: Aspose.Cells PDF export comments footnotes | C# PrintCommentsType.PrintSheetEnd | Excel comments to PDF footnotes | Aspose.Cells PdfSaveOptions example | Add page footer Aspose.Cells
// Common Searches: Aspose.Cells export comments as footnotes PDF | C# print Excel comments at end of sheet | How to show cell comments in PDF using Aspose.Cells | Add footnotes from Excel comments in PDF | Set page footer when converting Excel to PDF Aspose
// Developer Intent: Export an Excel worksheet to PDF where each cell comment appears as a footnote on the generated PDF pages.
// Use Cases: Product catalog with explanatory notes displayed as footnotes in the PDF brochure. | Compliance report where regulatory remarks stored as comments become footnotes after conversion. | Invoice PDF that includes terms and conditions as footnotes derived from cell comments.
// AI Prompts: Show how to place each comment on the same PDF page as its referenced cell instead of at the sheet end. | Provide code to customize the font, size, and color of comment footnotes when saving to PDF. | Explain how to extract comment text and insert it into a custom PDF footer using Aspose.Pdf after export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, add cell comments, configure PrintCommentsType.PrintSheetEnd to render those comments as footnotes, optionally set a page footer, and save the worksheet as a PDF with PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1.20);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(0.80);

        // Add comments that will be used as footnotes
        int commentIndex = sheet.Comments.Add("A2");
        sheet.Comments[commentIndex].Note = "Fresh apples sourced locally.";
        commentIndex = sheet.Comments.Add("A3");
        sheet.Comments[commentIndex].Note = "Ripe bananas imported from Ecuador.";

        // Configure the worksheet to print comments at the end of the sheet
        // This places the comments after the data, effectively acting as footnotes
        sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

        // Optional: add a page footer showing page numbers
        sheet.PageSetup.SetFooter(2, "&P of &N");

        // Save the workbook as PDF using default PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("CommentsFootnotes.pdf", pdfOptions);
    }
}
