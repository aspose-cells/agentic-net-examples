// Title: Generate a PDF report with cross‑sheet navigation hyperlinks using Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook with a Summary and Details worksheet, insert hyperlinks that jump to the opposite sheet, calculate formulas, and save the workbook as a PDF while keeping the links active using Aspose.Cells. | Configure PdfSaveOptions.ExportDocumentStructure to true so that hyperlinks are retained in the PDF output. | Add worksheet‑level hyperlinks via the Hyperlinks collection and verify they work after PDF conversion.
// Common Searches: asp.net aspose.cells add hyperlink to another worksheet and export to pdf | c# preserve Excel worksheet hyperlinks when saving as PDF with Aspose.Cells | how to create cross‑sheet navigation links in a PDF generated from Excel using Aspose.Cells | Aspose.Cells PdfSaveOptions ExportDocumentStructure example in C# | generate PDF report with clickable sheet links using Aspose.Cells for .NET
// Tags: add worksheet hyperlink Aspose.Cells C# | export workbook to PDF with active links Aspose.Cells | PdfSaveOptions ExportDocumentStructure C# | cross‑sheet navigation links Excel PDF Aspose | calculate formulas before PDF export Aspose.Cells

using System;
using Aspose.Cells;

// The example builds a workbook with Summary and Details sheets, adds hyperlinks that navigate between them, calculates any formulas, sets PdfSaveOptions.ExportDocumentStructure to retain links, and saves the workbook as a PDF where the cross‑sheet hyperlinks remain functional.
class PdfReportWithHyperlinks
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ---------- Sheet 1 : Summary ----------
        Worksheet summarySheet = workbook.Worksheets[0];
        summarySheet.Name = "Summary";

        // Add some content
        summarySheet.Cells["A1"].PutValue("Click to view Details");

        // Add a hyperlink that points to cell A1 of the Details sheet
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
        summarySheet.Hyperlinks.Add(0, 0, 1, 1, "Details!A1");

        // ---------- Sheet 2 : Details ----------
        int detailsIndex = workbook.Worksheets.Add();
        Worksheet detailsSheet = workbook.Worksheets[detailsIndex];
        detailsSheet.Name = "Details";

        // Add content to the Details sheet
        detailsSheet.Cells["A1"].PutValue("Here are the detailed data.");
        detailsSheet.Cells["A2"].PutValue("Back to Summary");

        // Add a hyperlink that points back to cell A1 of the Summary sheet
        detailsSheet.Hyperlinks.Add(1, 0, 1, 1, "Summary!A1");

        // Ensure any formulas are calculated before saving
        workbook.CalculateFormula();

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export document structure so that hyperlinks are retained in the PDF
            ExportDocumentStructure = true
        };

        // Save the workbook as a PDF file; hyperlinks will be embedded in the PDF
        workbook.Save("Report.pdf", pdfOptions);
    }
}
