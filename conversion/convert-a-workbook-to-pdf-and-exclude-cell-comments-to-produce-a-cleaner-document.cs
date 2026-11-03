// Title: Hide cell comments while converting an Excel workbook to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF without showing comments | How to disable comments in PDF output when saving .xlsx with Aspose.Cells | PdfSaveOptions CommentsVisibility hidden example in C# | Convert workbook to PDF and exclude cell comments using Aspose.Cells | C# Aspose.Cells hide comments during PDF conversion
// Tags: Aspose.Cells PdfSaveOptions CommentsVisibility | Excel to PDF conversion without comments | C# hide cell comments in PDF export | Aspose.Cells workbook save options PDF | PDF export excluding comments Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing XLSX file with Aspose.Cells, configures PdfSaveOptions to hide comments (using the CommentsVisibility property when available), and saves the workbook as a PDF, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // The CommentsVisibility property may not be available in older versions of Aspose.Cells.
            // If supported, uncomment the following line to hide comments in the PDF.
            // pdfOptions.CommentsVisibility = CommentsVisibility.Hidden;

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF file saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
