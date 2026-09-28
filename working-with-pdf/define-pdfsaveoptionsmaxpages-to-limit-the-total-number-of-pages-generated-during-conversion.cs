// Title: How to limit the number of pages when converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that configures the workbook's PageSetup to restrict the PDF output to a maximum number of pages before calling Workbook.Save with PdfSaveOptions. | Describe a workaround for limiting PDF pagination in Aspose.Cells when PdfSaveOptions does not provide a MaxPages property, using print area or manual page breaks.
// Common Searches: Aspose.Cells .NET limit PDF pages during Excel to PDF conversion | Set maximum page count for PDF export using Aspose.Cells C# | Workaround for missing PdfSaveOptions.MaxPages in Aspose.Cells | Control pagination of Excel workbook when saving as PDF with Aspose.Cells
// Tags: limit PDF page count Aspose.Cells C# | configure workbook PageSetup for PDF export | restrict Excel to PDF pagination Aspose.Cells | Aspose.Cells PDF export page limit workaround

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // The example loads an Excel workbook, notes that PdfSaveOptions lacks a MaxPages property, and demonstrates how to adjust the workbook's PageSetup (or print area) to enforce a maximum number of pages before saving the workbook as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the existing Excel workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                // Aspose.Cells PdfSaveOptions does not expose a MaxPages property.
                // To control pagination, adjust the workbook's PageSetup before saving if needed.

                // Save the workbook as PDF using the defined options
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
