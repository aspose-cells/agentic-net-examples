// Title: Validate single‑page PDF per worksheet using AllColumnsInOnePagePerSheet in Aspose.Cells for .NET
// Description: Creates a workbook with two sheets, each filled with 200 columns, saves it as PDF with PdfSaveOptions.AllColumnsInOnePagePerSheet = true and OnePagePerSheet = true, then uses WorkbookRender and SheetRender to confirm that every sheet renders exactly one page.
// Keywords: Aspose.Cells | AllColumnsInOnePagePerSheet | OnePagePerSheet | PDF pagination | WorkbookRender | SheetRender | C# example | page count verification | single page per sheet | Aspose.Cells PDF
// Common Searches: Aspose.Cells single page per worksheet PDF | AllColumnsInOnePagePerSheet option usage | How to check PDF page count with Aspose.Cells | C# verify PDF pagination Aspose.Cells | OnePagePerSheet PDF Aspose.Cells example
// Developer Intent: Ensure that a PDF generated from a workbook using AllColumnsInOnePagePerSheet and OnePagePerSheet settings contains exactly one page for each worksheet.
// Use Cases: Produce compact PDF reports where each sheet must fit on one page | Automated testing of PDF pagination settings in CI/CD pipelines | Validate workbook layout before distribution to clients | Create printable PDFs from wide tables without manual scaling | Integrate pagination verification into document generation services
// AI Prompts: Generate C# code that asserts each worksheet renders one PDF page when AllColumnsInOnePagePerSheet is enabled. | Describe how WorkbookRender calculates page counts with AllColumnsInOnePagePerSheet and OnePagePerSheet. | Provide alternative methods to verify PDF pagination without rendering each sheet. | Explain performance considerations when using SheetRender for page count verification. | Show how to log page count results for multiple worksheets in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook with two sheets, each filled with 200 columns, saves it as PDF with PdfSaveOptions.AllColumnsInOnePagePerSheet = true and OnePagePerSheet = true, then uses WorkbookRender and SheetRender to confirm that every sheet renders exactly one page.
class VerifyAllColumnsOnePagePerSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate each sheet with a large number of columns to force pagination
            for (int col = 0; col < 200; col++)
            {
                sheet1.Cells[0, col].PutValue($"Column {col + 1}");
                sheet2.Cells[0, col].PutValue($"Column {col + 1}");
            }

            // PDF save options to fit all columns on a single page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true,
                OnePagePerSheet = true
            };

            // Save the workbook as PDF (demonstrates actual file creation)
            workbook.Save("AllColumnsOnePagePerSheet.pdf", pdfOptions);

            // Rendering options (used for page count calculation)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                AllColumnsInOnePagePerSheet = true,
                OnePagePerSheet = true
                // ImageFormat is not required for page count calculation
            };

            // Use WorkbookRender to obtain the total page count of the rendered PDF
            WorkbookRender workbookRender = new WorkbookRender(workbook, renderOptions);
            int totalPageCount = workbookRender.PageCount;
            Console.WriteLine($"Total pages in PDF (rendered): {totalPageCount}");

            // Verify that each worksheet renders exactly one page
            bool verificationPassed = true;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                SheetRender sheetRender = new SheetRender(workbook.Worksheets[i], renderOptions);
                int sheetPageCount = sheetRender.PageCount;
                Console.WriteLine($"Worksheet '{workbook.Worksheets[i].Name}' page count: {sheetPageCount}");

                if (sheetPageCount != 1)
                {
                    verificationPassed = false;
                }
            }

            Console.WriteLine(verificationPassed
                ? "Verification succeeded: each worksheet has exactly one page."
                : "Verification failed: one or more worksheets have more than one page.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
