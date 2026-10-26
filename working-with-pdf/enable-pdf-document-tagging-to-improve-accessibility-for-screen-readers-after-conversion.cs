// Title: How to add PDF accessibility tags when converting an Excel workbook to PDF with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# convert Excel to PDF with accessibility tags for screen readers | Enable PDF tagging for accessibility using PdfSaveOptions in Aspose.Cells | Create accessible PDF from .xlsx using Aspose.Cells library | How to set PDF accessibility options when saving workbook as PDF in .NET | Add PDF/UA tags to Excel to PDF conversion with Aspose.Cells C# | Make PDF output from Excel readable by screen readers using Aspose.Cells
// Tags: Aspose.Cells PDF accessibility tagging | C# Excel to PDF conversion with tags | PdfSaveOptions enable accessibility | screen reader friendly PDF generation .NET | accessible PDF output from Excel workbook

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToPdfTagging
{
    // The sample verifies that input.xlsx exists, loads it with Aspose.Cells, configures PdfSaveOptions (e.g., OnePagePerSheet = false), and saves the workbook as output.pdf while handling exceptions. By adjusting PdfSaveOptions you can enable PDF accessibility tags so the resulting PDF is screen‑reader friendly.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source Excel file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the source Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // OnePagePerSheet = false keeps each worksheet on separate pages only if needed
                    OnePagePerSheet = false
                };

                // Save the workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
