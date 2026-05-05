using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class ConvertToPdfA1a
    {
        public static void Run()
        {
            // Source Excel file path
            string sourcePath = "input.xlsx";

            // Destination PDF/A‑1a file path
            string outputPath = "output_pdfa1a.pdf";

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options for PDF/A‑1a compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a,          // Set compliance level
                ExportDocumentStructure = true            // Optional: retain document structure
            };

            // Save the workbook as a PDF/A‑1a document
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine("Workbook successfully converted to PDF/A‑1a.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertToPdfA1a.Run();
        }
    }
}