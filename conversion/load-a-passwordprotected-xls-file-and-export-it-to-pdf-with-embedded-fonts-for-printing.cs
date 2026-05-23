using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            string sourceFile = "protected.xls";
            string outputPdf = "output.pdf";

            // Verify that the source file exists
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            try
            {
                // Load the password‑protected workbook
                var loadOpts = new LoadOptions
                {
                    Password = "excelPassword"
                };
                var workbook = new Workbook(sourceFile, loadOpts);

                // Configure PDF save options (fonts will be embedded by default where supported)
                var pdfOpts = new PdfSaveOptions();

                // Save the workbook as PDF
                workbook.Save(outputPdf, pdfOpts);

                Console.WriteLine("Password‑protected XLS has been converted to PDF.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}