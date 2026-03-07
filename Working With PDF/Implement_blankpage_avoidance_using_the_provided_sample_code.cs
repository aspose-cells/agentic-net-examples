using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class BlankPageAvoidanceDemo
    {
        public static void Run()
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Clear any existing content
            sheet.Cells.Clear();

            // Ensure there is at least one printable cell (even if empty) to avoid the "nothing to print" exception
            sheet.Cells["A1"].PutValue(string.Empty);
            sheet.PageSetup.PrintArea = "A1";

            // Configure PDF save options to suppress blank page output
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OutputBlankPageWhenNothingToPrint = false
            };

            // Define output file path
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Workbook_NoBlankPage.pdf");

            // Save the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            // OPTIONAL: Demonstrate the opposite behavior (blank page generated)
            pdfOptions.OutputBlankPageWhenNothingToPrint = true;
            string outputPathWithBlank = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Workbook_WithBlankPage.pdf");
            workbook.Save(outputPathWithBlank, pdfOptions);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            BlankPageAvoidanceDemo.Run();
        }
    }
}