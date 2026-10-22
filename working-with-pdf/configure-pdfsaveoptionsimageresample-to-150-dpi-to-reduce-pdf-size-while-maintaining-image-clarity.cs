// Title: How to set PdfSaveOptions.ImageResample to 150 DPI in Aspose.Cells for C# to shrink PDF size while keeping image clarity
// AI Prompts: Generate C# code that configures PdfSaveOptions.ImageResample to 150 DPI before calling Workbook.Save to produce a smaller PDF. | Modify the provided Aspose.Cells example to resample embedded images at 150 DPI during PDF export. | Explain the steps to apply a 150‑DPI image resampling setting in Aspose.Cells when converting an Excel workbook to PDF using C#.
// Common Searches: Aspose.Cells C# set PDF image resample DPI to 150 for smaller file size | How to reduce PDF output size by resampling images with PdfSaveOptions in Aspose.Cells | C# Aspose.Cells PDF export image DPI adjustment example | PdfSaveOptions ImageResample property usage in .NET
// Tags: Aspose.Cells PDF image resampling DPI | PdfSaveOptions ImageResample setting | reduce PDF size Aspose.Cells | C# export Excel to PDF with image DPI control | optimize PDF output image quality Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Loads an Excel workbook, creates PdfSaveOptions, and saves it as PDF. To lower the PDF size while preserving image clarity, set pdfOptions.ImageResample = 150 before calling workbook.Save.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default settings)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
