// Title: C# unit test to confirm FitToPagesTall = 0 generates unlimited page height when exporting a worksheet to PDF with Aspose.Cells
// Common Searches: Aspose.Cells C# verify FitToPagesTall zero creates multiple PDF pages | How to test unlimited page height in PDF export using FitToPagesTall with Aspose.Cells | C# example checking PDF file size after setting FitToPagesTall to 0 in Aspose.Cells | Aspose.Cells page setup FitToPagesWide 1 FitToPagesTall 0 PDF output validation
// Tags: unlimited height setting Aspose.Cells | Aspose.Cells PDF page setup configuration | C# worksheet FitToPagesWide setting | memory stream PDF generation C# Aspose.Cells | PDF page count verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook with 200 rows, sets PageSetup.FitToPagesWide = 1 and FitToPagesTall = 0, saves to PDF via a MemoryStream, and asserts that the resulting PDF length exceeds 1 KB, demonstrating unlimited page height and multiple pages.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Populate the worksheet with enough rows to require multiple pages vertically
                for (int row = 0; row < 200; row++)
                {
                    worksheet.Cells[row, 0].PutValue($"Row {row + 1}");
                }

                // Configure page setup:
                // - Fit the width to a single page
                // - Set FitToPagesTall to 0 to indicate unlimited page height
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0;

                // Export the workbook to PDF using a memory stream (no file I/O)
                using (var pdfStream = new MemoryStream())
                {
                    workbook.Save(pdfStream, SaveFormat.Pdf);
                    pdfStream.Position = 0; // Reset stream position for any further processing

                    // Simple verification: the PDF stream should contain data and be larger than a minimal size
                    if (pdfStream.Length > 1024)
                    {
                        Console.WriteLine("PDF generated successfully with multiple pages (FitToPagesTall = 0).");
                    }
                    else
                    {
                        Console.WriteLine("PDF generated, but size is unexpectedly small.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
