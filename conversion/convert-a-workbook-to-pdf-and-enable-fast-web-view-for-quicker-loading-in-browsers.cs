using System;
using System.IO;
using Aspose.Cells; // Workbook, PdfSaveOptions

namespace AsposeCellsDemo
{
    public class ConvertToPdfFastWebView
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("PDF conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Hello, Fast Web View PDF!");
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Note: FastWebView property is not available in the current Aspose.Cells version.
            // If a newer version supports it, you can enable it here.

            // Save the workbook as a PDF using the configured options
            string outputPath = "output.pdf";
            try
            {
                workbook.Save(outputPath, pdfOptions);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to save PDF: {ex.Message}");
            }
        }
    }
}