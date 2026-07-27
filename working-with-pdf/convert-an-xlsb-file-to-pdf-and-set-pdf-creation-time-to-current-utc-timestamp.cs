using System;
using Aspose.Cells;
using Aspose.Cells.Utility; // Aspose.Cells conversion utilities

// Author: Aspose.Cells .NET example – XLSB to PDF with UTC CreatedTime
class Program
{
    static void Main()
    {
        // Paths for source XLSB and target PDF
        string sourcePath = "input.xlsb";
        string outputPath = "output.pdf";

        // Load options (can be default for XLSB)
        LoadOptions loadOptions = new LoadOptions();

        // PDF save options – set creation time to current UTC
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CreatedTime = DateTime.UtcNow
        };

        // Perform conversion using Aspose.Cells utility
        ConversionUtility.Convert(sourcePath, loadOptions, outputPath, pdfOptions);
    }
}