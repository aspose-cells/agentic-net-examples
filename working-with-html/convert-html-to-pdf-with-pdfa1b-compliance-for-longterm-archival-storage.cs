using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class HtmlToPdfA1b
{
    static void Main()
    {
        // Source HTML file path
        string htmlPath = "input.html";

        // Destination PDF file path
        string pdfPath = "output.pdf";

        // Load the HTML file into a workbook
        Workbook workbook = new Workbook(htmlPath);

        // Create PDF save options and set compliance to PDF/A‑1b
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the workbook as a PDF with the specified compliance level
        workbook.Save(pdfPath, saveOptions);

        Console.WriteLine("HTML successfully converted to PDF/A‑1b.");
    }
}