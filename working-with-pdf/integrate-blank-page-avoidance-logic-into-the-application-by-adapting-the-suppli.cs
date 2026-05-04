using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBlankPageAvoidance
{
    public class BlankPageAvoidanceDemo
    {
        public static void Run()
        {
            // Create a new workbook with some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("First page data");
            sheet.Cells["A2"].PutValue("More data");

            // Simulate a second sheet that would be blank (nothing to print)
            Worksheet blankSheet = workbook.Worksheets.Add("BlankSheet");
            // Keep the sheet empty or hide it to represent no printable content
            blankSheet.IsVisible = false;

            // Configure PDF save options to avoid generating a blank page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not output a blank page when a sheet has nothing to print
                OutputBlankPageWhenNothingToPrint = false,

                // Additionally, ignore completely blank pages during pagination
                PrintingPageType = PrintingPageType.IgnoreBlank
            };

            // Define output path (desktop folder for demonstration)
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Workbook_NoBlankPages.pdf");

            // Save the workbook using the configured options
            workbook.Save(outputPath, pdfOptions);

            // Inform the user
            Console.WriteLine($"Workbook saved without blank pages to: {outputPath}");
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