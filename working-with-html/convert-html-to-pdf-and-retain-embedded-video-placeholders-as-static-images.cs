using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

class HtmlToPdfWithVideoPlaceholders
{
    static void Main()
    {
        // Path to the source HTML file that may contain embedded video players.
        string htmlFilePath = "input.html";

        // Path for the resulting PDF file.
        string pdfFilePath = "output.pdf";

        // Load the HTML file into a Workbook.
        // Aspose.Cells parses the HTML and creates corresponding worksheets,
        // cells, images and web extensions (e.g., video players).
        Workbook workbook = new Workbook(htmlFilePath);

        // Create PDF save options.
        // The default behavior renders web extensions (such as video players)
        // as static image placeholders in the PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF. Video players will appear as images.
        workbook.Save(pdfFilePath, pdfOptions);

        Console.WriteLine("HTML has been converted to PDF with video placeholders rendered as static images.");
    }
}