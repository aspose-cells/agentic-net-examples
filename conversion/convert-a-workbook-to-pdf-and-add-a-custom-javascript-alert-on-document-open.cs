using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and add some sample data
            // ------------------------------------------------------------
            var workbook = new Workbook();                     // create workbook
            var worksheet = workbook.Worksheets[0];            // get first worksheet
            worksheet.Cells["A1"].PutValue("Sample Data");     // add data
            worksheet.Cells["B1"].PutValue(12345);             // add more data

            // ------------------------------------------------------------
            // 2. Save the workbook as PDF
            // ------------------------------------------------------------
            string outputPdf = "sample_with_js.pdf";

            // Configure PDF save options (JavaScript embedding not supported in this version)
            var pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPdf));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save directly to the final PDF file
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine("PDF created successfully.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}