using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to demonstrate the conversion
        workbook.Worksheets[0].Cells["A1"].PutValue("Converted to PDF");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set the creation time of the PDF to the current processing time
            CreatedTime = DateTime.Now
        };

        // Define the output PDF file name
        string outputPdf = "ConvertedDocument.pdf";

        // Save the workbook as PDF using the specified options
        workbook.Save(outputPdf, pdfOptions);

        // Inform the user
        Console.WriteLine($"PDF saved to '{outputPdf}' with CreatedTime = {pdfOptions.CreatedTime}");
    }
}