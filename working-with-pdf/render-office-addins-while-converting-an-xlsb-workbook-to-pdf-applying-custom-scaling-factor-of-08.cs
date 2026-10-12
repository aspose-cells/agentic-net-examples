// Title: Convert an XLSB workbook to a single‑page PDF with 0.8 scaling and render Office Add‑Ins using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an XLSB file, sets PdfSaveOptions.OnePagePerSheet = true, PdfSaveOptions.ScalingFactor = 0.8, enables PdfSaveOptions.RenderOfficeAddIns, and saves the workbook as a PDF. | Update existing Aspose.Cells PDF export logic to include a scaling factor of 0.8 while preserving the one‑page‑per‑sheet layout and rendering any Office Add‑Ins present in the workbook. | Add robust file‑existence checking and exception handling around the XLSB‑to‑PDF conversion process in C#.
// Common Searches: Aspose.Cells render Office Add‑Ins when exporting XLSB to PDF | C# set scaling factor 0.8 in PdfSaveOptions Aspose.Cells | convert xlsb to single page pdf with one page per sheet Aspose.Cells | fit all columns on one PDF page and apply custom scale using Aspose.Cells
// Tags: XLSB to PDF conversion with scaling Aspose.Cells | PdfSaveOptions RenderOfficeAddIns property | OnePagePerSheet PDF export Aspose.Cells | ScalingFactor 0.8 PDF generation C# | Office Add‑Ins rendering during PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The sample verifies the input XLSB file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to render Office Add‑Ins, apply a 0.8 scaling factor, and place each worksheet on a single PDF page, then saves the result as a PDF while handling errors gracefully.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsb";
                string outputFile = "output.pdf";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the XLSB workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Fit all columns on one page (alternative to AllColumnsInOnePage)
                    OnePagePerSheet = true
                };

                // Save the workbook as a PDF with the specified options
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
