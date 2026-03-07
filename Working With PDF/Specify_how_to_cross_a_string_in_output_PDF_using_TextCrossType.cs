using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a long text that will exceed the cell width
        worksheet.Cells["A1"].PutValue("This is a very long text that will cross cell boundaries when saved to PDF");

        // Reduce the column width to force the text to overflow
        worksheet.Cells.SetColumnWidth(0, 5);

        // Create PDF save options and set the TextCrossType
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Options: Default, CrossKeep, CrossOverride, StrictInCell
        pdfOptions.TextCrossType = TextCrossType.CrossKeep;

        // Save the workbook as PDF with the specified TextCrossType
        workbook.Save("CrossStringDemo.pdf", pdfOptions);
    }
}