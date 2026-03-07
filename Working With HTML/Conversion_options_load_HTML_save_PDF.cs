using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Create HTML load options
            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions
            {
                SupportDivTag = true,
                AutoFitColsAndRows = true
            };

            // Load the HTML file into a workbook using the load options
            Workbook workbook = new Workbook("input.html", htmlLoadOptions);

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                EmbedStandardWindowsFonts = true
            };

            // Save the workbook as a PDF file using the save options
            workbook.Save("output.pdf", pdfSaveOptions);
        }
    }
}