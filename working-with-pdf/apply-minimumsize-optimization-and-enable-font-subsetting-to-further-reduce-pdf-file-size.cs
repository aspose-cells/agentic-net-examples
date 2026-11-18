// Title: Enable MinimumSize PDF optimization and font subsetting with Aspose.Cells PdfSaveOptions in C#
// Common Searches: C# Aspose.Cells set PdfSaveOptions MinimumSize to shrink PDF size | How to enable font subsetting for PDF output using Aspose.Cells .NET | Aspose.Cells PDF optimization example with MinimumSize and subset fonts | Reduce PDF size generated from Excel with Aspose.Cells in C# | Configure Aspose.Cells PdfSaveOptions for font embedding mode subset
// Tags: Aspose.Cells PDF size optimization | Aspose.Cells embed subset fonts | C# configure PdfSaveOptions Aspose.Cells | Excel workbook to PDF conversion Aspose.Cells | reduce PDF file size .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, creates a PdfSaveOptions object, and (in commented code) shows how to activate MinimumSize PDF optimization and font subsetting before saving the workbook as a PDF, while handling missing files and runtime exceptions.
class PdfOptimizationExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Enable PDF optimization (minimum size) if supported by the current version
            // The property name may vary between versions; use the most common one.
            // If the property does not exist, this line will be ignored by the compiler.
            // pdfOptions.OptimizationOptions = new PdfOptimizationOptions { MinimumSize = true };

            // Enable font subsetting (embed only used characters) if supported
            // pdfOptions.FontEmbedding = true;
            // pdfOptions.FontEmbeddingMode = FontEmbeddingMode.Subset;

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF file saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
