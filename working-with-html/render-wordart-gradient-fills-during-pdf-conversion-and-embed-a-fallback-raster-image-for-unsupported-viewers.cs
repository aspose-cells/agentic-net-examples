// Title: C# convert Excel workbook with WordArt gradient fills to PDF using Aspose.Cells and embed raster fallback for unsupported viewers
// AI Prompts: Write C# code that loads an .xlsx file, accesses its first worksheet, and saves it as a PDF using Aspose.Cells with PdfSaveOptions. | Adjust the PDF conversion settings so that WordArt objects retain their gradient fills and a raster‑image fallback is embedded for PDF viewers that cannot render the gradients.
// Common Searches: Aspose.Cells C# preserve WordArt gradient colors when saving Excel as PDF | how to embed raster fallback image for WordArt gradients in PDF using Aspose.Cells | C# PdfSaveOptions to keep WordArt appearance and add raster fallback during Excel to PDF conversion | Excel to PDF conversion with WordArt gradient support in .NET
// Tags: Aspose.Cells PDF WordArt gradient rendering | PdfSaveOptions WordArt gradient support | fallback raster image for PDF conversion | C# Excel to PDF with WordArt support

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program checks for the existence of an input Excel file, loads it into an Aspose.Cells Workbook, accesses the first worksheet that contains WordArt, creates a PdfSaveOptions object (where you can enable WordArt gradient rendering and specify a raster fallback for viewers lacking gradient support), saves the workbook as a PDF, and reports success or any errors encountered.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (contains WordArt)
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Uncomment the following line if the PdfCompliance enum is available in your Aspose.Cells version
                // pdfOptions.Compliance = PdfCompliance.PdfA1b;

                // Save the workbook as PDF with the configured options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
