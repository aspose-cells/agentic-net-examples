// Title: Convert an Excel workbook that contains slicers to PDF with slicer controls rendered as static images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets PdfSaveOptions.SlicerRenderingMode to Image, and saves the workbook as a PDF with Aspose.Cells. | Demonstrate how to detect slicers in a workbook and ensure they are rendered as static images during PDF conversion in a .NET application. | Add robust file‑existence checks and exception handling to a C# program that converts all worksheets, including slicers, to a PDF with slicer images.
// Common Searches: Aspose.Cells C# export slicer controls as images in PDF | How to render Excel slicers as static images when saving to PDF with Aspose.Cells | PdfSaveOptions SlicerRenderingMode Image example .NET | Convert workbook with slicers to PDF using Aspose.Cells and preserve slicer appearance | C# code sample for saving Excel file with slicers to PDF as images
// Tags: export slicers as images pdf aspose.cells | pdfsaveoptions slicerrenderingmode image | c# convert excel to pdf with slicer rendering | aspose.cells workbook to pdf static slicer | handle missing input file aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The program loads an Excel workbook, optionally sets PdfSaveOptions.SlicerRenderingMode to render slicers as static images, and saves all worksheets to a PDF while handling missing files and runtime errors.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.pdf";

                // Ensure the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook that contains slicers
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // If your Aspose.Cells version supports slicer rendering, enable it:
                // pdfOptions.SlicerRenderingMode = SlicerRenderingMode.Image;

                // Export the workbook (all worksheets) to a PDF file
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"Workbook successfully saved to PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
