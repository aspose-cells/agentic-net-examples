using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XlsxToPdfConverter
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the output PDF file
            string pdfPath = "output.pdf";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.Error.WriteLine($"Source file not found: '{sourcePath}'");
                return;
            }

            try
            {
                // Load the workbook from the specified file path
                Workbook workbook = new Workbook(sourcePath);

                // Save the loaded workbook as a PDF document
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{pdfPath}'");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}