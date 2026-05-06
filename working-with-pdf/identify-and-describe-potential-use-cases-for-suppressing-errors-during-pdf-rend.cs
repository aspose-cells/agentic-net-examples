using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates when and why you might want to suppress rendering errors (IgnoreError = true)
    public class SuppressRenderingErrorsDemo
    {
        public static void Run()
        {
            // Create a new workbook and add data that could cause rendering problems
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Normal data
            sheet.Cells["A1"].PutValue("Normal Text");
            sheet.Cells["A2"].PutValue(12345);

            // Add a shape with an unsupported image format (e.g., a corrupted image)
            // This could throw an error during PDF conversion if not suppressed
            try
            {
                // Intentionally using an invalid image path to simulate a rendering error
                sheet.Pictures.Add(2, 0, "nonexistent_image.xyz");
            }
            catch
            {
                // Ignored – the purpose is to have a shape that will cause an error later
            }

            // Add a chart with missing data series (another source of rendering errors)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
            // Do not add any series to the chart – this can cause a rendering exception

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Enabling IgnoreError hides these errors and continues rendering the rest of the document
                IgnoreError = true
            };

            // Save the workbook as PDF; any rendering errors will be suppressed
            workbook.Save("SuppressErrorsOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with rendering errors suppressed (IgnoreError = true).");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SuppressRenderingErrorsDemo.Run();
        }
    }
}