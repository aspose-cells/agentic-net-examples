using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data to demonstrate the effect
        for (int col = 0; col < 30; col++)
        {
            sheet.Cells[0, col].PutValue("Header " + (col + 1));
            sheet.Cells[1, col].PutValue("Sample data " + (col + 1));
        }

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Fit all columns of each sheet onto a single PDF page
        pdfOptions.AllColumnsInOnePagePerSheet = true;

        // Save the workbook as PDF using the configured options
        workbook.Save("AllColumnsOnePage.pdf", pdfOptions);
    }
}