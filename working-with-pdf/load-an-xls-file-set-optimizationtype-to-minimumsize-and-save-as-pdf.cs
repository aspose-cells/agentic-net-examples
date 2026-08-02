using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Author: Aspose.Cells .NET example – load XLS, set PDF optimization, save as PDF

        // Define source and destination file paths
        string sourceFile = "input.xls";
        string destinationFile = "output.pdf";

        // Create load options (default settings)
        LoadOptions loadOptions = new LoadOptions();

        // Configure PDF save options with MinimumSize optimization
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OptimizationType = PdfOptimizationType.MinimumSize
        };

        // Convert the XLS workbook to PDF using the provided ConversionUtility rule
        ConversionUtility.Convert(sourceFile, loadOptions, destinationFile, pdfOptions);
    }
}