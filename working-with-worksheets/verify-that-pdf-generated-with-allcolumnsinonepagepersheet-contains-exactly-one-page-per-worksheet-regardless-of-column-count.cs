// Title: Aspose.Cells for .NET – Verify PDF with AllColumnsInOnePagePerSheet renders one page per worksheet
// Description: Creates a workbook with two sheets, each filled with 200 columns, saves it to PDF using PdfSaveOptions (AllColumnsInOnePagePerSheet = true, OnePagePerSheet = true), renders the same workbook with ImageOrPrintOptions, retrieves the page count via WorkbookRender, and checks that the count matches the number of worksheets, confirming a single‑page‑per‑sheet PDF.
// Keywords: Aspose.Cells | AllColumnsInOnePagePerSheet | OnePagePerSheet | PDF pagination | C# | WorkbookRender page count | verify single page per sheet | Aspose.Cells PDF options | render workbook to PDF | page count validation
// Common Searches: Aspose.Cells PDF one page per worksheet | AllColumnsInOnePagePerSheet example C# | how to get PDF page count with Aspose.Cells | verify PDF pagination Aspose.Cells | C# render workbook to single page PDF
// Developer Intent: Ensure that setting AllColumnsInOnePagePerSheet (and OnePagePerSheet) produces exactly one PDF page for each worksheet, regardless of column width.
// Use Cases: Automated testing of PDF layout for reports with many columns | Generating printable PDFs where each sheet must occupy a single page | Validating pagination settings before deploying document generation pipelines
// AI Prompts: Generate a C# unit test that asserts WorkbookRender.PageCount equals Workbook.Worksheets.Count when AllColumnsInOnePagePerSheet is enabled. | Explain how PdfSaveOptions and ImageOrPrintOptions work together to force a single PDF page per worksheet in Aspose.Cells. | Provide code to log a pass/fail message after comparing rendered page count with worksheet count for multi‑sheet workbooks.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook with two sheets, each filled with 200 columns, saves it to PDF using PdfSaveOptions (AllColumnsInOnePagePerSheet = true, OnePagePerSheet = true), renders the same workbook with ImageOrPrintOptions, retrieves the page count via WorkbookRender, and checks that the count matches the number of worksheets, confirming a single‑page‑per‑sheet PDF.
class VerifyAllColumnsOnePagePerSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add a second worksheet
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Sheet2");

            // Populate each worksheet with a large number of columns
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int col = 0; col < 200; col++)
                {
                    sheet.Cells[0, col].PutValue($"Column {col + 1}");
                    sheet.Cells[1, col].PutValue($"Sample data {col + 1}");
                }
            }

            // Configure PDF save options to force all columns onto a single page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true,
                OnePagePerSheet = true
            };

            // Save the workbook as PDF
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, pdfOptions);

            // Use ImageOrPrintOptions for rendering to obtain page counts
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                AllColumnsInOnePagePerSheet = true,
                OnePagePerSheet = true
            };

            // Render the workbook and get the total page count
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
            int totalPages = renderer.PageCount;
            int expectedPages = workbook.Worksheets.Count;

            // Output verification result
            Console.WriteLine($"Total pages rendered: {totalPages}");
            Console.WriteLine($"Expected pages (one per worksheet): {expectedPages}");
            Console.WriteLine(totalPages == expectedPages
                ? "Verification passed: each worksheet is rendered on exactly one page."
                : "Verification failed: page count does not match the number of worksheets.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
