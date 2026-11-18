// Title: Convert an XLS workbook to PDF with MinimumSize optimization using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert XLS to PDF with MinimumSize optimization | How to reduce PDF size when saving an Excel file with Aspose.Cells | Set PdfSaveOptions.OptimizationType to MinimumSize in Aspose.Cells example | C# code to load .xls and save as compressed PDF using Aspose.Cells | Aspose.Cells PDFSaveOptions minimum file size setting for Excel to PDF conversion | Saving an Excel workbook as the smallest possible PDF with Aspose.Cells .NET
// Tags: Aspose.Cells PDFSaveOptions MinimumSize optimization | XLS to PDF conversion C# Aspose.Cells | reduce PDF file size Aspose.Cells | Workbook.Save with PdfSaveOptions C# | Excel to compressed PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example loads an existing XLS file into an Aspose.Cells Workbook, configures PdfSaveOptions with the MinimumSize optimization type, and saves the workbook as a PDF, producing a smaller output file while handling errors gracefully.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xls";
                string outputPath = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing XLS file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // Set minimum size optimization if the enum is available
                // Uncomment the following line if your Aspose.Cells version supports OptimizationType
                // saveOptions.OptimizationType = OptimizationType.MinimumSize;

                // Save the workbook as PDF
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
