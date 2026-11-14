// Title: Convert an Excel workbook containing slicers to PDF while preserving slicer rendering using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to export an Excel file that contains slicers to PDF while keeping the slicer visuals intact | preserving slicer appearance when saving a workbook as PDF using Aspose.Cells PdfSaveOptions | example code for checking if an Excel file exists before converting it to PDF with Aspose.Cells in .NET | adding custom document properties to a PDF generated from an Excel workbook after Aspose.Cells conversion | why Aspose.Cells does not transfer Excel custom properties to PDF and how to handle it in C#
// Tags: Aspose.Cells slicer PDF export | PdfSaveOptions default slicer rendering | C# verify Excel file existence before conversion | post‑generation PDF metadata injection for Aspose.Cells output | preserve slicer layout during Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input Excel file exists, loads it with Aspose.Cells, applies the default PdfSaveOptions (which automatically render slicers) to save the workbook as a PDF, and explains that custom document properties are not transferred to the PDF and must be added with a separate PDF library. Errors are caught and reported.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains slicers.
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Aspose.Cells renders slicers automatically when saving to PDF,
            // so no additional settings are required beyond the default behavior.

            // NOTE: Aspose.Cells PdfSaveOptions does not expose a direct API to
            // transfer custom document properties to the PDF. If needed, such
            // metadata can be added after PDF generation using a PDF library.

            // Save the workbook as a PDF file with the specified options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
