// Title: Convert Excel workbook to PDF while embedding external linked images using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert Excel to PDF embed linked images | How to include external pictures in PDF generated from Excel with Aspose.Cells | C# save workbook as PDF with linked image embedding using Aspose.Cells | PdfSaveOptions default behavior embed linked images Aspose.Cells .NET | Convert .xlsx to .pdf and embed external images using Aspose.Cells API
// Tags: Aspose.Cells PDFSaveOptions embed linked images | Excel to PDF conversion with external image embedding | C# Aspose.Cells workbook to PDF with pictures | PDF export linked graphics Aspose.Cells .NET | Aspose.Cells external image handling in PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, configures PdfSaveOptions (which by default embeds any linked images), and saves the workbook as a PDF, ensuring that external images referenced in the worksheet are embedded directly in the resulting PDF file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    var tempWb = new Workbook();
                    tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    tempWb.Save(inputPath);
                }

                // Load the source workbook.
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default behavior embeds linked images).
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF using the specified options.
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
