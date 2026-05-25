using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExcelToPdfWithPageBreaks
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can iterate through all worksheets if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Ensure the worksheet has the desired page breaks.
        // Example: add a horizontal page break after row 20 (cell A21)
        sheet.HorizontalPageBreaks.Add("A21");
        // Example: add a vertical page break after column 5 (cell F1)
        sheet.VerticalPageBreaks.Add("F1");
        // -------------------------------------------------

        // Optional: retrieve automatic page breaks based on current print settings
        ImageOrPrintOptions printOpts = new ImageOrPrintOptions();
        CellArea[] automaticBreaks = sheet.GetPrintingPageBreaks(printOpts);
        // The 'automaticBreaks' array can be inspected or used for custom logic.

        // Configure PDF save options to keep the original pagination.
        // By default Aspose.Cells respects manual page breaks when converting to PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = false;               // Do not force a single page per sheet
        pdfOptions.AllColumnsInOnePagePerSheet = false;   // Keep column pagination as defined

        // Save the workbook as PDF, preserving the page breaks.
        workbook.Save("output.pdf", pdfOptions);
    }
}