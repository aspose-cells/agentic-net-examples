using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Load HTML content into the workbook (replace with your actual HTML file path)
            Workbook workbook = new Workbook("input.html");

            // Optional: calculate any formulas before saving
            workbook.CalculateFormula();

            // Configure PDF save options (e.g., export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as a PDF file using the specified options
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}