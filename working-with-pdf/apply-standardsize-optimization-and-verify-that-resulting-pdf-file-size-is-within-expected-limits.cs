using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Author: Aspose.Cells .NET example - Standard PDF optimization and size verification
class PdfStandardOptimizationDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF Standard Optimization Example");
        sheet.Cells["A2"].PutValue("This PDF should prioritize print quality over file size.");

        // Configure PDF save options with Standard optimization (high print quality)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OptimizationType = PdfOptimizationType.Standard
        };

        // Define output file path
        string outputPath = "StandardOptimized.pdf";

        // Save the workbook as PDF using the specified options
        workbook.Save(outputPath, pdfOptions);

        // Verify the resulting PDF file size
        const long expectedMaxSizeBytes = 500_000; // example threshold (≈500 KB)
        FileInfo pdfInfo = new FileInfo(outputPath);
        long actualSize = pdfInfo.Length;

        Console.WriteLine($"PDF saved to '{outputPath}'. Size: {actualSize} bytes.");

        if (actualSize <= expectedMaxSizeBytes)
        {
            Console.WriteLine("File size is within the expected limits.");
        }
        else
        {
            Console.WriteLine($"Warning: File size exceeds the expected limit of {expectedMaxSizeBytes} bytes.");
        }
    }
}