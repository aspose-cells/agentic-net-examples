using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Desired PDF output path
        string pdfPath = "output.pdf";

        // Load the workbook from the Excel file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(sourcePath);

        // Create PDF save options and set the creation timestamp for audit tracking
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CreatedTime = DateTime.Now
        };

        // Save the workbook as PDF with the specified options (uses Workbook.Save(string, SaveOptions))
        workbook.Save(pdfPath, pdfOptions);
    }
}