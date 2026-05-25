using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSM workbook
            string sourcePath = "input.xlsm";

            // Path for the resulting PDF/A‑2a file
            string destPath = "output.pdf";

            // Load the XLSM workbook (lifecycle rule: create/load)
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options for PDF/A‑2a compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the PDF/A‑2a compliance level
                Compliance = PdfCompliance.PdfA2a,

                // Ensure Windows fonts are embedded (helps preserve appearance)
                EmbedStandardWindowsFonts = true,

                // Calculate formulas before saving (optional but common)
                CalculateFormula = true
            };

            // Save the workbook as PDF/A‑2a (lifecycle rule: save with options)
            workbook.Save(destPath, pdfOptions);

            Console.WriteLine("Conversion completed: " + destPath);
        }
    }
}