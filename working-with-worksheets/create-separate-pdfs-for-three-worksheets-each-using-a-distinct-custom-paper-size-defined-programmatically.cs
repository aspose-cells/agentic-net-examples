// Title: Export Each Worksheet to a Separate PDF with Custom Paper Sizes using Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook with three sheets, assign a distinct custom page size (in inches) to each sheet via PageSetup.CustomPaperSize, and generate an individual PDF per sheet by configuring PdfSaveOptions.SheetSet to render only the chosen worksheet.
// Keywords: Aspose.Cells | C# | .NET PDF export | custom page dimensions | custom paper size | PdfSaveOptions | SheetSet | PageSetup.CustomPaperSize | separate worksheet PDF | programmatic PDF generation | Aspose.Cells US | Aspose.Cells Europe | Aspose.Cells Asia
// Common Searches: Aspose.Cells set custom paper size inches | Export single worksheet to PDF with Aspose.Cells | Create PDF per sheet with different page size C# | How to use SheetSet to save one worksheet as PDF | C# Aspose.Cells custom page size example
// Developer Intent: Produce three PDF files, each containing one worksheet that uses its own custom paper size.
// Use Cases: Print labels of varying dimensions from separate sheets in a single workbook. | Generate marketing flyers or brochures where each sheet requires a unique page size. | Export individual financial statements with tailored layouts for archiving. | Create region‑specific reports (e.g., US, EU, APAC) with different page formats in one project.
// AI Prompts: Convert the custom paper size definitions from inches to centimeters in the provided Aspose.Cells code. | Add a header and footer to each PDF while preserving the distinct custom page sizes. | Show a one‑line approach to batch‑convert all worksheets to PDFs with their custom sizes, eliminating the helper method.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomPaperSizePdf
{
    // Shows how to build a workbook with three sheets, assign a distinct custom page size (in inches) to each sheet via PageSetup.CustomPaperSize, and generate an individual PDF per sheet by configuring PdfSaveOptions.SheetSet to render only the chosen worksheet.
    class Program
    {
        static void Main()
        {
            // -------------------- Create a new workbook --------------------
            Workbook workbook = new Workbook();

            // -------------------- Prepare three worksheets --------------------
            // Worksheet 0 (default)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("Data for Sheet 1");

            // Worksheet 1
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Data for Sheet 2");

            // Worksheet 2
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
            sheet3.Cells["A1"].PutValue("Data for Sheet 3");

            // -------------------- Set distinct custom paper sizes (in inches) --------------------
            // Sheet1: 2" x 2"
            sheet1.PageSetup.PaperSize = PaperSizeType.Custom;          // Enable custom size
            sheet1.PageSetup.CustomPaperSize(2.0, 2.0);

            // Sheet2: 3" x 4"
            sheet2.PageSetup.PaperSize = PaperSizeType.Custom;
            sheet2.PageSetup.CustomPaperSize(3.0, 4.0);

            // Sheet3: 5" x 7"
            sheet3.PageSetup.PaperSize = PaperSizeType.Custom;
            sheet3.PageSetup.CustomPaperSize(5.0, 7.0);

            // -------------------- Save each worksheet as a separate PDF --------------------
            // Helper method to save a single sheet to PDF
            void SaveSheetToPdf(int sheetIndex, string fileName)
            {
                // Configure PDF save options to render only the specified sheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // SheetSet selects the sheet by its zero‑based index
                    SheetSet = new SheetSet(new int[] { sheetIndex })
                };

                // Save the workbook; only the selected sheet will be rendered
                workbook.Save(fileName, pdfOptions);
            }

            // Save each sheet with its own PDF file
            SaveSheetToPdf(0, "Sheet1_CustomSize.pdf");
            SaveSheetToPdf(1, "Sheet2_CustomSize.pdf");
            SaveSheetToPdf(2, "Sheet3_CustomSize.pdf");

            Console.WriteLine("PDF files created successfully.");
        }
    }
}
