using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCommentFootnotePdf
{
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

            // Add comments to cells that we want to appear as footnotes
            int commentIndex = sheet.Comments.Add("A2");
            Comment commentA2 = sheet.Comments[commentIndex];
            commentA2.Note = "Freshly picked apples.";

            commentIndex = sheet.Comments.Add("B3");
            Comment commentB3 = sheet.Comments[commentIndex];
            commentB3.Note = "Discounted price for bananas.";

            // Configure the page setup to print comments at the end of the sheet (footnotes)
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

            // Create PDF save options (default options are sufficient for comments)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure that the document structure is exported so footnotes are retained
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF; comments will appear as footnotes
            workbook.Save("WorkbookWithComments.pdf", pdfOptions);

            Console.WriteLine("PDF generated with comments as footnotes.");
        }
    }
}