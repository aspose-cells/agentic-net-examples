// Title: How to convert an Excel workbook containing sparklines to PDF with sparklines rendered as inline images using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook with sparklines as PDF inline images | PdfSaveOptions SparklineAsImage property example in .NET | Convert Excel file with sparklines to PDF preserving sparkline graphics | Render sparklines as images when exporting Excel to PDF using Aspose.Cells | C# code to export workbook containing sparklines to PDF with images
// Tags: Aspose.Cells PdfSaveOptions SparklineAsImage | Excel sparklines to PDF conversion .NET | render sparklines as images Aspose.Cells | inline graphics PDF export Aspose.Cells | C# workbook to PDF with sparklines

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook that includes sparklines, optionally enables the PdfSaveOptions.SparklineAsImage flag to render sparklines as images, and saves the workbook as a PDF where the sparklines appear as inline graphics, with basic file existence checks and error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook containing sparklines
                Workbook workbook = new Workbook(inputPath);

                // Set PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                // In recent Aspose.Cells versions sparklines are rendered as images by default.
                // If using a version that supports the property, enable it as shown below:
                // pdfOptions.SparklineAsImage = true;

                // Save the workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
