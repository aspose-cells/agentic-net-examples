using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfVersionCompatibilityDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF version compatibility demo");

        // Create PDF save options and set the compliance level to PDF 1.4
        PdfSaveOptions options = new PdfSaveOptions
        {
            Compliance = PdfCompliance.Pdf14   // Enables compatibility with older PDF readers
        };

        // Save the workbook as a PDF using the specified options
        workbook.Save("Output_Pdf14.pdf", options);

        Console.WriteLine("PDF saved with PDF 1.4 compliance.");
    }
}