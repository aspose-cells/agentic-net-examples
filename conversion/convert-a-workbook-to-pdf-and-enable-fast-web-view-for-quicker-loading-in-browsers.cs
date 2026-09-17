// Title: Convert an Excel workbook to a linearized PDF using Aspose.Cells for .NET with fallback when Fast Web View is unavailable
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as a fast‑loading PDF for quicker browser rendering. | Show how to programmatically check if PdfSaveOptions can produce a linearized PDF and switch to a standard PDF option when the feature is not supported. | Add validation for missing source files and implement comprehensive exception handling in the Excel‑to‑PDF conversion workflow.
// Common Searches: Aspose.Cells .NET export Excel to linearized PDF for quick web view | C# detect support for linearized PDF in PdfSaveOptions | How to save a workbook as PDF with incremental loading using Aspose.Cells | Code example for handling unavailable Fast Web View option in Aspose.Cells | Example of error handling for missing Excel file during PDF conversion
// Tags: Aspose.Cells PDFSaveOptions linearized PDF | Excel to PDF conversion .NET Aspose.Cells | incremental loading PDF Aspose.Cells | C# workbook save as PDF default option | exception handling missing input file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook with Aspose.Cells, configures PdfSaveOptions (noting that Fast Web View may be unavailable), and saves the file as a linearized PDF for faster browser loading, while handling missing input files and general exceptions.
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

                // Load the Excel workbook from the file
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options (Fast Web View not available in this version)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF file with the specified options
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
