// Title: C# – Export an Aspose.Cells Worksheet to PDF with a 5×7 in Custom Paper Size (Save Method)
// Description: This C# example creates a workbook, fills cells with sample data, sets the worksheet’s page size to 5 in × 7 in using PageSetup.CustomPaperSize, configures PdfSaveOptions, and saves the workbook as a PDF via Workbook.Save.
// Keywords: Aspose.Cells | C# PDF export | custom paper size | PageSetup.CustomPaperSize | PdfSaveOptions | 5x7 inch PDF | worksheet to PDF | non‑standard page size | Aspose.Cells .NET | Workbook.Save
// Common Searches: Aspose.Cells set custom paper size C# | Export worksheet to PDF with 5x7 inches Aspose.Cells | Workbook.Save custom page dimensions PDF | How to use PageSetup.CustomPaperSize in Aspose.Cells | PdfSaveOptions custom size example
// Developer Intent: Create a PDF of a worksheet using a 5 in × 7 in custom page size.
// Use Cases: Printing photo‑size reports on 5×7 paper | Generating receipts that match small envelope dimensions | Producing catalog cards or label sheets with exact page size | Creating printable flyers with non‑standard dimensions
// AI Prompts: Write C# code that sets a worksheet's custom paper size to 5 inches by 7 inches and saves it as a PDF using Aspose.Cells. | Show how to apply PdfSaveOptions while preserving a custom page size in Aspose.Cells .NET. | Explain how to verify the page dimensions of the PDF generated from Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomPaperPdf
{
    // This C# example creates a workbook, fills cells with sample data, sets the worksheet’s page size to 5 in × 7 in using PageSetup.CustomPaperSize, configures PdfSaveOptions, and saves the workbook as a PDF via Workbook.Save.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherries");
            sheet.Cells["B4"].PutValue(15);

            // Set custom paper size (width: 5 inches, height: 7 inches)
            sheet.PageSetup.CustomPaperSize(5.0, 7.0);

            // Create PDF save options (optional: you can set additional options here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the Save method with options
            workbook.Save("CustomPaperSizeOutput.pdf", pdfOptions);

            Console.WriteLine("Worksheet rendered to PDF with custom paper dimensions successfully.");
        }
    }
}
