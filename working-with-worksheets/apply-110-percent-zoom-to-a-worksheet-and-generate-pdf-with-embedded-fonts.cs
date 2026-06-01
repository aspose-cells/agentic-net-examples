using System;
using Aspose.Cells;

namespace AsposeCellsZoomPdfDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (optional, just to have content in the PDF)
                worksheet.Cells["A1"].PutValue("Zoom 110% and Embedded Fonts Demo");
                worksheet.Cells["A2"].PutValue(DateTime.Now);

                // Apply 110 percent zoom to the worksheet (percent scale between 10 and 400)
                worksheet.Zoom = 110;

                // Configure PDF save options to embed fonts
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Embed standard Windows TrueType fonts (required for embedding)
                    EmbedStandardWindowsFonts = true,

                    // Optional: specify a default font to use for Unicode characters
                    DefaultFont = "Arial"
                    // FontEncoding defaults to Identity, which works for most cases
                };

                // Save the workbook as a PDF with the specified options
                workbook.Save("Zoom110_EmbeddedFonts.pdf", pdfOptions);

                Console.WriteLine("PDF generated with 110% zoom and embedded fonts.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}