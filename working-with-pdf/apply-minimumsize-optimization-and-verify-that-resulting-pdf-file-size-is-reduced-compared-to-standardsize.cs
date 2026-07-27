using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfOptimizationComparison
{
    static void Main()
    {
        // Create a workbook with sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Aspose.Cells PDF Optimization Example");
        worksheet.Cells["A2"].PutValue("Comparing Standard vs MinimumSize PDF sizes.");

        // Save PDF with Standard optimization (default high quality)
        string standardPath = "StandardPdf.pdf";
        PdfSaveOptions standardOptions = new PdfSaveOptions
        {
            OptimizationType = PdfOptimizationType.Standard
        };
        workbook.Save(standardPath, standardOptions);

        // Save PDF with MinimumSize optimization (prioritizes smaller file size)
        string minSizePath = "MinimumSizePdf.pdf";
        PdfSaveOptions minSizeOptions = new PdfSaveOptions
        {
            OptimizationType = PdfOptimizationType.MinimumSize
        };
        workbook.Save(minSizePath, minSizeOptions);

        // Retrieve file sizes for comparison
        long standardSize = new FileInfo(standardPath).Length;
        long minSize = new FileInfo(minSizePath).Length;

        Console.WriteLine($"Standard PDF size: {standardSize} bytes");
        Console.WriteLine($"MinimumSize PDF size: {minSize} bytes");

        // Verify that MinimumSize produces a smaller file
        if (minSize < standardSize)
        {
            Console.WriteLine("MinimumSize optimization reduced the PDF file size.");
        }
        else
        {
            Console.WriteLine("MinimumSize optimization did not reduce the PDF file size.");
        }
    }
}

// Author: Aspose.Cells .NET example demonstrating PDF optimization types.