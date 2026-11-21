// Title: Export specific worksheets to a single PDF using PdfSaveOptions.SheetSet in C# with Aspose.Cells
// Common Searches: how to export only certain worksheets to one PDF file using Aspose.Cells C# | Aspose.Cells PdfSaveOptions SheetSet example with worksheet indices | C# code to save selected Excel sheets as a combined PDF using Aspose.Cells | set zero‑based sheet indices in PdfSaveOptions for PDF export Aspose.Cells | export non‑adjacent Excel sheets to PDF with Aspose.Cells C#
// Tags: PdfSaveOptions SheetSet configuration | selected worksheets PDF export Aspose.Cells | worksheet index based PDF conversion C# | single PDF from multiple Excel sheets Aspose.Cells | zero‑based sheet indices PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, configures PdfSaveOptions with a SheetSet containing zero‑based worksheet indices (e.g., 0, 2, 4), and saves those selected sheets together as a single PDF file using Aspose.Cells in C#.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "selected_sheets.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Specify the worksheets to export (zero‑based indices)
                    SheetSet = new SheetSet(new int[] { 0, 2, 4 })
                };

                // Save the selected sheets as a single PDF file
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
