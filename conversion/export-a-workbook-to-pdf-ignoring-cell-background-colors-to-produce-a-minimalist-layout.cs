// Title: Export an Excel workbook to PDF without cell background colors using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF ignoring cell fill colors | How to remove background shading when saving workbook as PDF with Aspose.Cells | PdfSaveOptions IgnoreCellBackground example in C# | Generate a clean PDF from an Excel file using Aspose.Cells .NET | Set Aspose.Cells PDF options to hide cell backgrounds
// Tags: Aspose.Cells PDF export ignore cell background | PdfSaveOptions IgnoreCellBackground C# | Excel to PDF conversion without shading Aspose | minimalist PDF layout Aspose.Cells | reflection set PDF options Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The code loads an Excel workbook, creates a PdfSaveOptions object, uses reflection to enable the IgnoreCellBackground property when available, and saves the workbook as a PDF, producing a PDF that omits all cell background colors for a minimalist appearance.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set IgnoreCellBackground if the property exists in the current version
            var ignoreProp = typeof(PdfSaveOptions).GetProperty("IgnoreCellBackground");
            if (ignoreProp != null && ignoreProp.CanWrite)
            {
                ignoreProp.SetValue(pdfOptions, true);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
