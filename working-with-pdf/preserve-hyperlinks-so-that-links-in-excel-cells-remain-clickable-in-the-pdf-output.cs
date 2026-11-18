// Title: How to keep Excel cell hyperlinks clickable when converting to PDF with Aspose.Cells for .NET (C#)
// Common Searches: how to retain Excel cell hyperlinks when exporting to PDF using Aspose.Cells C# | Aspose.Cells PDF export keep hyperlinks enabled | C# convert .xlsx to .pdf preserving clickable links | enable hyperlink preservation in Aspose.Cells PDFSaveOptions | why are hyperlinks missing in PDF generated from Excel with Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions hyperlink preservation | C# Excel to PDF conversion with active links | IsHyperlinkEnabled property Aspose.Cells | preserve Excel cell hyperlinks in PDF output | XLSX to PDF conversion with active hyperlinks

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExample
{
    // The example loads an XLSX workbook, configures PdfSaveOptions (noting that hyperlinks are preserved by default or can be explicitly enabled via IsHyperlinkEnabled), and saves the workbook as a PDF while ensuring any cell hyperlinks remain clickable. It also includes basic file existence checks and exception handling.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Hyperlinks are preserved by default in recent Aspose.Cells versions.
                // If using a version that provides the property, enable it as shown below:
                // pdfOptions.IsHyperlinkEnabled = true;

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display an informative message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
