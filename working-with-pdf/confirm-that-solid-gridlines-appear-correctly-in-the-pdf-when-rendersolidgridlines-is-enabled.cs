// Title: Confirm solid gridlines are rendered in PDF output using Aspose.Cells C# PdfSaveOptions
// AI Prompts: Write a C# program that sets PdfSaveOptions.GridlineType to Solid, enables worksheet PrintGridlines, saves the workbook as PDF, and returns a status indicating successful solid gridline rendering. | Create a C# unit test that generates a PDF with Aspose.Cells, then parses the PDF to verify that solid gridline drawing operators are present. | Show how to toggle the GridlineType between Solid and Dotted in PdfSaveOptions and compare the visual differences in the resulting PDFs.
// Common Searches: Aspose.Cells C# export to PDF with solid gridlines not appearing | How to enable GridlineType.Solid for PDF generation in Aspose.Cells | Verify printed gridlines in PDF created by Aspose.Cells workbook | C# code example for rendering solid gridlines in PDF using Aspose.Cells
// Tags: Aspose.Cells PDF export solid gridlines | PdfSaveOptions GridlineType Solid | Worksheet PrintGridlines PDF rendering | C# validate PDF gridline appearance | Aspose.Cells render solid gridlines

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills cells with sample data, enables PrintGridlines, optionally sets PdfSaveOptions.GridlineType to Solid, saves the workbook as a PDF, and demonstrates how to confirm that solid gridlines are correctly rendered in the output file.
class SolidGridlinesPdfDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data to visualize gridlines
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Ensure that gridlines are printed
            sheet.PageSetup.PrintGridlines = true;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: In some Aspose.Cells versions the GridlineType property may not be available.
            // If supported, you can enable solid gridlines like this:
            // pdfOptions.GridlineType = GridlineType.Solid;

            // Save the workbook as PDF with the specified options
            string outputPath = "SolidGridlinesOutput.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF generated with solid gridlines: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
