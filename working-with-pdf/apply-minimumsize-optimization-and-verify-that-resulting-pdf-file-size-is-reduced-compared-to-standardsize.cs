using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with a sizable amount of data to make the PDF noticeable
        for (int row = 0; row < 500; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define file names for the two PDF outputs
        string standardPdfPath = "StandardOptimization.pdf";
        string minimumPdfPath = "MinimumSizeOptimization.pdf";

        // -----------------------------------------------------------------
        // Save with the default (Standard) optimization
        // -----------------------------------------------------------------
        // No special options are needed; the default OptimizationType is Standard
        workbook.Save(standardPdfPath, SaveFormat.Pdf);

        // Get the file size of the Standard optimized PDF
        long standardSize = new FileInfo(standardPdfPath).Length;
        Console.WriteLine($"Standard optimization PDF size: {standardSize} bytes");

        // -----------------------------------------------------------------
        // Save with MinimumSize optimization
        // -----------------------------------------------------------------
        PdfSaveOptions minSizeOptions = new PdfSaveOptions();
        minSizeOptions.OptimizationType = PdfOptimizationType.MinimumSize; // Apply MinimumSize optimization

        workbook.Save(minimumPdfPath, minSizeOptions);

        // Get the file size of the MinimumSize optimized PDF
        long minimumSize = new FileInfo(minimumPdfPath).Length;
        Console.WriteLine($"MinimumSize optimization PDF size: {minimumSize} bytes");

        // -----------------------------------------------------------------
        // Verify that the MinimumSize PDF is smaller than the Standard PDF
        // -----------------------------------------------------------------
        if (minimumSize < standardSize)
        {
            Console.WriteLine("Verification passed: MinimumSize PDF is smaller than Standard PDF.");
        }
        else
        {
            Console.WriteLine("Verification failed: MinimumSize PDF is not smaller than Standard PDF.");
        }
    }
}