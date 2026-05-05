using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class ConvertWorkbookToPdfA1a
    {
        public static void Run()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path for the resulting PDF/A‑1a file
            string destPath = "output_pdfa1a.pdf";

            // Load the workbook from the source file
            Workbook workbook = new Workbook(sourcePath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the compliance level to PDF/A‑1a
                Compliance = PdfCompliance.PdfA1a
            };

            // Save the workbook as PDF/A‑1a using the save options
            workbook.Save(destPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF/A‑1a: {destPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertWorkbookToPdfA1a.Run();
        }
    }
}