using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Demo PDF with maximum compression");

        // Configure PDF save options for maximum compression
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // Flate compression provides the highest compression for PDF content
        saveOptions.PdfCompression = PdfCompressionCore.Flate;
        // Further reduce size by optimizing for minimum file size
        saveOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output_max_compression.pdf", saveOptions);
    }
}

// Author: Example demonstrating maximum PDF compression using Aspose.Cells.