using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that references external CSS files
            string htmlFilePath = @"C:\Input\sample.html";

            // Path where the resulting PDF will be saved
            string pdfOutputPath = @"C:\Output\sample.pdf";

            try
            {
                // Verify that the input HTML file exists
                if (!File.Exists(htmlFilePath))
                {
                    Console.WriteLine($"Input file not found: {htmlFilePath}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfOutputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the HTML file into a new Workbook instance
                Workbook workbook = new Workbook(htmlFilePath);

                // Save the workbook as PDF, preserving the visual appearance defined by the original CSS
                workbook.Save(pdfOutputPath, SaveFormat.Pdf);

                Console.WriteLine("HTML converted to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}