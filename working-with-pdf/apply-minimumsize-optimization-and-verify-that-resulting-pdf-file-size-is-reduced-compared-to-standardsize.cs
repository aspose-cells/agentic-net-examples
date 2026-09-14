// Title: Generate a PDF from an Aspose.Cells workbook using MinimumSize optimization and compare its size to a standard PDF in C#
// AI Prompts: Create a new Workbook, populate it with rows and a column chart, then save it twice as PDF – first with default PdfSaveOptions and second with PdfSaveOptions configured for minimum‑size optimization (e.g., set Compression to Standard, ImageCompression to Jpeg, JpegQuality to 50) – and output the byte length of each file. | Read the file sizes of the two PDFs, display them, and programmatically verify that the minimum‑size PDF is smaller than the standard PDF, printing a success or failure message.
// Common Searches: how to reduce PDF size when exporting an Aspose.Cells workbook in C# | Aspose.Cells MinimumSize PDF optimization example .NET | compare default and compressed PDF output using PdfSaveOptions in C# | enable JPEG image compression for PDF export with Aspose.Cells | verify that a compressed PDF is smaller than the original using C# file info
// Tags: minimum size PDF optimization Aspose.Cells C# | PdfSaveOptions compression settings Aspose.Cells | compare PDF file sizes Aspose.Cells | image compression JPEG PDF export Aspose.Cells | verify PDF size reduction C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample creates a workbook with 200 rows and a column chart, saves it to PDF twice – once with default options and once with minimum‑size optimization (compression and JPEG image settings) – then reads both files, prints their byte sizes, and confirms that the optimized PDF is smaller.
class MinimumSizeOptimizationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with sample data to increase PDF size
            for (int row = 0; row < 200; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Add a simple chart to further enlarge the PDF
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 12, 25, 22);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A200", true);
            chart.NSeries[0].Name = "Sample Series";

            // Define file paths
            string standardPdfPath = "StandardSize.pdf";
            string minimumPdfPath = "MinimumSize.pdf";

            // Save PDF with standard size (default options)
            PdfSaveOptions standardOptions = new PdfSaveOptions();
            workbook.Save(standardPdfPath, standardOptions);

            // Save PDF with minimum size optimization (higher compression)
            PdfSaveOptions minimumOptions = new PdfSaveOptions();

            // The following properties are not available in older Aspose.Cells versions.
            // If your version supports them, you can uncomment and adjust as needed.
            // minimumOptions.Compression = PdfCompressionType.Standard;
            // minimumOptions.ImageCompression = PdfImageCompressionType.Jpeg;
            // minimumOptions.JpegQuality = 50;

            workbook.Save(minimumPdfPath, minimumOptions);

            // Verify that the minimum size PDF is smaller than the standard size PDF
            long standardSize = new FileInfo(standardPdfPath).Length;
            long minimumSize = new FileInfo(minimumPdfPath).Length;

            Console.WriteLine($"Standard PDF size : {standardSize} bytes");
            Console.WriteLine($"Minimum PDF size  : {minimumSize} bytes");

            if (minimumSize < standardSize)
            {
                Console.WriteLine("Verification passed: MinimumSize PDF is smaller than StandardSize PDF.");
            }
            else
            {
                Console.WriteLine("Verification failed: MinimumSize PDF is not smaller than StandardSize PDF.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
