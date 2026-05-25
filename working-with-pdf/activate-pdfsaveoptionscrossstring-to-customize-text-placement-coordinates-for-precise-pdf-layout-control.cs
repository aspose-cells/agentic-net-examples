using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfCrossStringDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a long text that will exceed the cell width
        worksheet.Cells["A1"].PutValue("This is a very long text that will exceed the width of the cell and demonstrate TextCrossType behavior.");
        worksheet.Cells["B1"].PutValue("Short text");

        // Narrow the first column to force the text to cross cell boundaries
        worksheet.Cells.SetColumnWidth(0, 5);

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Activate the TextCrossType (CrossString equivalent) to keep the overflowing text visible
        pdfOptions.TextCrossType = TextCrossType.CrossKeep;

        // Optional: set a default font to ensure proper rendering
        pdfOptions.DefaultFont = "Arial";

        // Save the workbook as a PDF with the configured options
        workbook.Save("CrossStringDemo.pdf", pdfOptions);
    }
}