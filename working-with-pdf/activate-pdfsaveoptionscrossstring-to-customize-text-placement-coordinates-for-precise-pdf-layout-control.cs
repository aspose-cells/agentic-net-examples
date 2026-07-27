using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text for PDF layout");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // NOTE: The documentation does not define a 'CrossString' property on PdfSaveOptions.
        // If such a property exists in a newer version, set it here.
        // pdfOptions.CrossString = ...; // <-- placeholder for missing API evidence

        // Save the workbook as PDF using the options
        workbook.Save("Output.pdf", pdfOptions);
    }
}