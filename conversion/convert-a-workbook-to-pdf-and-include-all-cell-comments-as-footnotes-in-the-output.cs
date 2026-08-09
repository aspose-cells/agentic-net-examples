// Title: Export Excel workbook to PDF with cell comments as footnotes using Aspose.Cells (C#)
// Description: Learn how to convert an Aspose.Cells workbook to PDF while rendering every cell comment as a footnote. The example sets PageSetup.PrintComments to PrintSheetEnd, configures PdfSaveOptions (ExportDocumentStructure, CalculateFormula), and saves the file as a PDF with comments listed at the end of each sheet.
// Keywords: Aspose.Cells PDF export | C# Excel to PDF comments footnote | PrintCommentsType.PrintSheetEnd | PdfSaveOptions ExportDocumentStructure | Excel comments footnotes Aspose | Workbook.Save PDF with comments | Aspose.Cells C# tutorial
// Common Searches: Aspose.Cells export PDF with comments as footnotes | C# print Excel cell comments at end of sheet PDF | PdfSaveOptions PrintCommentsType example | How to include Excel comments in PDF using Aspose | Convert workbook to PDF footnote comments
// Developer Intent: Create a PDF from an Excel workbook where all cell comments are displayed as footnotes at the end of each worksheet.
// Use Cases: Generate a product catalog PDF that shows item notes stored in cell comments as footnotes. | Produce a financial statement PDF with analyst remarks from comments rendered as end‑sheet footnotes. | Automate a compliance report PDF where regulatory notes are kept in comments and appear as footnotes.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as PDF and render cell comments as footnotes. | Explain the role of PageSetup.PrintComments = PrintSheetEnd when exporting to PDF. | Provide step‑by‑step instructions to configure PdfSaveOptions for exporting comments as footnotes.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCommentFootnotePdf
{
    // Learn how to convert an Aspose.Cells workbook to PDF while rendering every cell comment as a footnote. The example sets PageSetup.PrintComments to PrintSheetEnd, configures PdfSaveOptions (ExportDocumentStructure, CalculateFormula), and saves the file as a PDF with comments listed at the end of each sheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Add comments to cells that will become footnotes
            int commentIndex = sheet.Comments.Add("A2");
            Comment commentA2 = sheet.Comments[commentIndex];
            commentA2.Note = "Fresh apples from the orchard.";

            commentIndex = sheet.Comments.Add("A3");
            Comment commentA3 = sheet.Comments[commentIndex];
            commentA3.Note = "Ripe bananas imported from Ecuador.";

            // Configure the page setup to print comments at the end of the sheet
            // This makes comments appear as footnotes in the PDF output
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

            // Create PDF save options (lifecycle: create)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure that the document structure is retained (optional but useful)
                ExportDocumentStructure = true,
                // Calculate formulas before saving (good practice)
                CalculateFormula = true
            };

            // Save the workbook as PDF (lifecycle: save)
            workbook.Save("WorkbookWithCommentsFootnotes.pdf", pdfOptions);

            Console.WriteLine("Workbook saved to PDF with comments rendered as footnotes.");
        }
    }
}
