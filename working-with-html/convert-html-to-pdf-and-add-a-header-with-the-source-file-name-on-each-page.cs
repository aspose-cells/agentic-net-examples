using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;   // PdfSaveOptions

namespace AsposeCellsHtmlToPdfWithHeader
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path of the source HTML file
                string htmlPath = "source.html";

                // Verify that the HTML file exists before loading
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                    return;
                }

                // Load the HTML file into a workbook
                Workbook workbook = new Workbook(htmlPath);

                // Add a header to each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Left: file name (&F)
                    // Center: page number of total pages
                    // Right: current date
                    sheet.PageSetup.SetHeader(0, "&F");                     // Left
                    sheet.PageSetup.SetHeader(1, "Page &P of &N");          // Center
                    sheet.PageSetup.SetHeader(2, "&D");                    // Right
                }

                // Configure PDF save options (optional: display document title in the PDF window)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    DisplayDocTitle = true
                };

                // Save the workbook as a PDF file
                string pdfPath = "output.pdf";
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF with header on each page.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}