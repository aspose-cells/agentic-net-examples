using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SuppressRenderingErrorsDemo
    {
        public static void Run()
        {
            // Create a new workbook (you can also load an existing file)
            Workbook workbook = new Workbook();

            // Add some data that might cause rendering warnings (e.g., a missing font)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample text with a non‑existent font");
            Style style = workbook.CreateStyle();
            style.Font.Name = "NonExistentFont";
            sheet.Cells["A1"].SetStyle(style);

            // Configure PDF save options to hide rendering errors
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // When set to true, errors such as shape, image, or chart rendering issues are ignored
                IgnoreError = true
            };

            // Save the workbook as PDF using the configured options
            workbook.Save("SuppressedErrorsOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved successfully with rendering errors suppressed.");
        }

        public static void Main()
        {
            Run();
        }
    }
}