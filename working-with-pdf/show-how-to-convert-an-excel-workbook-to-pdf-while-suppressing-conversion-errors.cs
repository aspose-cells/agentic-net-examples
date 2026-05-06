using System;
using Aspose.Cells;

class ConvertExcelToPdf
{
    static void Main()
    {
        // Load the source Excel workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);   // LoadOptions not needed for default format

        // Create PDF save options and suppress conversion/rendering errors
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.IgnoreError = true;   // Hide errors such as shape, image, or chart rendering issues

        // Save the workbook as PDF using the configured options
        string destPath = "output.pdf";
        workbook.Save(destPath, pdfOptions);

        Console.WriteLine("Excel file has been converted to PDF with errors ignored.");
    }
}