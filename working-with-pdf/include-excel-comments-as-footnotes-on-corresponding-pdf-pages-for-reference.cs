using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCommentFootnotesPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(20);

            // Add comments to cells that will act as footnotes
            int commentIdx1 = sheet.Comments.Add("A2");
            Comment comment1 = sheet.Comments[commentIdx1];
            comment1.Note = "Fresh red apples from the orchard.";

            int commentIdx2 = sheet.Comments.Add("A3");
            Comment comment2 = sheet.Comments[commentIdx2];
            comment2.Note = "Ripe bananas imported from Ecuador.";

            // Configure the page setup to print comments at the end of the sheet
            // This makes the comments appear as footnotes in the generated PDF
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

            // Optional: add a footer that shows page numbers (useful for multi‑page PDFs)
            sheet.PageSetup.SetFooter(0, "Page &P of &N");

            // Create PDF save options (default options are sufficient for this scenario)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF; comments will be rendered as footnotes
            workbook.Save("CommentsAsFootnotes.pdf", pdfOptions);
        }
    }
}