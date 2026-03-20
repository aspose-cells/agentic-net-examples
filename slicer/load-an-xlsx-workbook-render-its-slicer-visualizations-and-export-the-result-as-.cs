using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsSlicerToPdf
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX workbook that contains slicers
            string sourcePath = "input.xlsx";

            // Load the workbook (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options (optional, e.g., export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF; slicer visualizations are rendered as part of the PDF
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook with slicers successfully exported to PDF: {outputPath}");
        }
    }
}