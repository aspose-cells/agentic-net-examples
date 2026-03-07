using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class OutputBlankPageDemo
    {
        public static void Run()
        {
            // Create a new workbook with a single worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet is visible (required by Aspose.Cells)
            sheet.IsVisible = true;

            // Add a dummy value to avoid "nothing to print" exception
            sheet.Cells["A1"].PutValue("Sample");

            // Initialize PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // -------------------------------------------------
            // Example 1: Do NOT output a blank page when nothing to print
            // -------------------------------------------------
            pdfOptions.OutputBlankPageWhenNothingToPrint = false;

            // Define output path for the first PDF
            string outputPath1 = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "NoBlankPage.pdf");

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath1, pdfOptions);

            // -------------------------------------------------
            // Example 2: Output a blank page when nothing to print
            // -------------------------------------------------
            pdfOptions.OutputBlankPageWhenNothingToPrint = true;

            // Define output path for the second PDF
            string outputPath2 = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "WithBlankPage.pdf");

            // Save the workbook again with the updated option
            workbook.Save(outputPath2, pdfOptions);
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}