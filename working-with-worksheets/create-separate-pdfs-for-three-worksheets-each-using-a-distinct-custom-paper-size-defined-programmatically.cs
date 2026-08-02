// Title: C# – Export each worksheet to a separate PDF with distinct custom paper sizes using Aspose.Cells
// Description: Creates a workbook with three worksheets, assigns unique custom paper dimensions (3×5, 4×6, 5×7 inches) via PageSetup.CustomPaperSize, and saves each sheet individually as a PDF by setting PdfSaveOptions.SheetSet to the appropriate sheet index.
// Keywords: Aspose.Cells | C# | custom paper size | PDF export | PageSetup.CustomPaperSize | PdfSaveOptions SheetSet | separate PDF per worksheet | programmatic page size | .NET | export worksheet to PDF
// Common Searches: set custom paper size for worksheet Aspose.Cells PDF | export each worksheet to its own PDF C# Aspose.Cells | use SheetSet to save selected worksheets as PDF | different page dimensions per sheet Aspose.Cells .NET | how to define custom page size in inches for PDF export
// Developer Intent: Generate individual PDF files for each worksheet, each using a programmatically defined custom paper size.
// Use Cases: Print reports where sections require different page formats | Create marketing handouts with varying dimensions per worksheet | Produce compliance documents that must follow specific page‑size standards
// AI Prompts: Write C# code that sets a custom paper size of 8.5x11 inches for a worksheet and saves it as a PDF with Aspose.Cells. | Explain how PdfSaveOptions.SheetSet can be used to export selected worksheets to separate PDFs with individual page setups. | Provide a C# loop that iterates through all worksheets, assigns a custom paper size in centimeters, and saves each as an individual PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomPaperSizePdf
{
    // Creates a workbook with three worksheets, assigns unique custom paper dimensions (3×5, 4×6, 5×7 inches) via PageSetup.CustomPaperSize, and saves each sheet individually as a PDF by setting PdfSaveOptions.SheetSet to the appropriate sheet index.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Populate each sheet with sample data
            sheet1.Cells["A1"].PutValue("Data for Sheet 1");
            sheet2.Cells["A1"].PutValue("Data for Sheet 2");
            sheet3.Cells["A1"].PutValue("Data for Sheet 3");

            // Define distinct custom paper sizes (width, height) in inches
            // Sheet1: 3 x 5 inches
            sheet1.PageSetup.CustomPaperSize(3.0, 5.0);
            // Sheet2: 4 x 6 inches
            sheet2.PageSetup.CustomPaperSize(4.0, 6.0);
            // Sheet3: 5 x 7 inches
            sheet3.PageSetup.CustomPaperSize(5.0, 7.0);

            // Save each worksheet as a separate PDF using SheetSet to select the sheet
            SaveWorksheetAsPdf(workbook, 0, "Sheet1_CustomSize.pdf");
            SaveWorksheetAsPdf(workbook, 1, "Sheet2_CustomSize.pdf");
            SaveWorksheetAsPdf(workbook, 2, "Sheet3_CustomSize.pdf");
        }

        // Helper method that creates PdfSaveOptions with a SheetSet containing a single sheet index
        static void SaveWorksheetAsPdf(Workbook wb, int sheetIndex, string outputFile)
        {
            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Use SheetSet constructor that accepts an array of sheet indexes
            pdfOptions.SheetSet = new SheetSet(new int[] { sheetIndex });

            // Save the workbook to PDF; only the specified sheet will be rendered
            wb.Save(outputFile, pdfOptions);
        }
    }
}
