// Title: Export an Excel workbook that contains slicers to a single-page PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel sheet with slicers to PDF on one page | How to keep slicers visible when saving a workbook as PDF using Aspose.Cells | PdfSaveOptions OnePagePerSheet slicer rendering example in Aspose.Cells | Save Excel workbook with slicers to PDF without page breaks in .NET | Aspose.Cells code to convert a workbook containing slicers to PDF
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | export workbook with slicers to PDF | render slicers in PDF Aspose.Cells | C# single-page PDF export Aspose.Cells | Excel slicer PDF conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an existing Excel file that includes slicers, configures PdfSaveOptions.OnePagePerSheet to true so the entire sheet and its slicers are rendered on a single PDF page, and saves the result as OutputWorkbook.pdf while handling missing files and runtime exceptions.
class ExportWorkbookWithSlicers
{
    static void Main()
    {
        try
        {
            // Path to the input workbook that contains slicers
            string inputPath = "InputWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found – {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to render the whole sheet (including slicers) on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Export the workbook to PDF; slicers will appear on the same page as the sheet content
            workbook.Save("OutputWorkbook.pdf", pdfOptions);

            Console.WriteLine("Workbook successfully exported to OutputWorkbook.pdf");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
