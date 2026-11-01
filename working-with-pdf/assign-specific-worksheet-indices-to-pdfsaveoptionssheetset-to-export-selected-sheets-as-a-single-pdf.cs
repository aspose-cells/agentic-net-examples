// Title: How to export selected worksheets (e.g., first and third) to a single PDF using Aspose.Cells PdfSaveOptions.SheetSet in C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, sets PdfSaveOptions.SheetSet to specific worksheet indices, and saves those sheets as one PDF. | Show how to configure Aspose.Cells PdfSaveOptions to include only worksheets 0 and 2 when converting to PDF in a .NET application. | Provide an example of using a SheetSet index array in PdfSaveOptions to combine selected Excel sheets into a single PDF document.
// Common Searches: Aspose.Cells C# export only certain worksheets to PDF using SheetSet | PdfSaveOptions SheetSet array example for selecting sheets in .NET | How to save first and third Excel sheets as one PDF with Aspose.Cells | Selective worksheet PDF conversion Aspose.Cells .NET tutorial | C# Aspose.Cells PDF export specific sheet indices
// Tags: Aspose.Cells PdfSaveOptions SheetSet selection | C# export specific worksheets to PDF | Aspose.Cells selective sheet PDF conversion | PdfSaveOptions worksheet index array | Aspose.Cells combine multiple sheets into single PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, creates a PdfSaveOptions object with SheetSet set to the indices of the first and third worksheets (0 and 2), and saves those selected sheets together as a single PDF file using Aspose.Cells in C#.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "selected_sheets.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options and specify sheets to include (first and third)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    SheetSet = new SheetSet(new int[] { 0, 2 })
                };

                // Save the selected sheets as a single PDF file
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
