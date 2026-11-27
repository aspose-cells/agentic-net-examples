// Title: Create a PDF from an Aspose.Cells workbook with each worksheet on a separate page using PdfSaveOptions in C#
// Common Searches: C# Aspose.Cells generate PDF where each Excel sheet becomes its own page | How to set OnePagePerSheet in PdfSaveOptions for workbook to PDF conversion | Aspose.Cells export multiple worksheets to single PDF with separate pages | Save Excel workbook as PDF with one page per sheet using Aspose.Cells API | PdfSaveOptions OnePagePerSheet example code in C#
// Tags: Aspose.Cells PdfSaveOptions per-sheet pagination | C# export Excel workbook to PDF per worksheet | separate PDF pages for each Excel sheet | Excel to PDF conversion with Aspose.Cells | configure PDF save options for sheet-level pagination

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook with two worksheets, enables PdfSaveOptions.OnePagePerSheet, and saves the workbook as a PDF where each worksheet is rendered on its own page.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Populate first worksheet with sample data
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].PutValue("Hello from Sheet1");

        // Add a second worksheet and populate it
        int sheetIndex = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheetIndex];
        sheet2.Name = "Sheet2";
        sheet2.Cells["A1"].PutValue("Hello from Sheet2");

        // Configure PDF save options to generate one page per worksheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true;

        // Export the workbook to PDF with the specified options
        workbook.Save("ExportedWorkbook.pdf", pdfOptions);
    }
}
