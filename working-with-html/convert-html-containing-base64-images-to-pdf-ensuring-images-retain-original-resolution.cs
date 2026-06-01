using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains base64‑encoded images
            string htmlPath = "input.html";

            // Load the HTML into a Workbook (Aspose.Cells can parse HTML as a spreadsheet)
            Workbook workbook = new Workbook(htmlPath);

            // Configure PDF save options to keep image resolution.
            // SetImageResample with a high PPI (e.g., 300) and maximum JPEG quality (100)
            // ensures that images are not down‑sampled during the conversion.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.SetImageResample(300, 100);

            // Save the workbook as PDF. The images embedded in the HTML will be rendered
            // at their original resolution because we prevented resampling.
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"HTML converted to PDF successfully: {pdfPath}");
        }
    }
}