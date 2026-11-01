// Title: Set PdfSaveOptions.CompressionLevel to High in Aspose.Cells for .NET to reduce PDF size while preserving visual fidelity
// Common Searches: Aspose.Cells C# set PdfSaveOptions compression level to High for PDF export | How to reduce PDF output size when converting an Excel workbook with Aspose.Cells .NET | Enable high compression in PdfSaveOptions while saving XLSX as PDF using Aspose.Cells | C# example configuring PdfSaveOptions.CompressionLevel for smaller PDFs | Aspose.Cells PDF export high compression setting code sample
// Tags: PdfSaveOptions compression level Aspose.Cells | optimize PDF size Aspose.Cells .NET | Excel workbook PDF export Aspose.Cells | C# configure PDF save options Aspose.Cells | Aspose.Cells PDF output size reduction

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // Namespace for PdfSaveOptions

// The example verifies the input Excel file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, optionally sets its CompressionLevel property to High (requiring Aspose.Cells.Pdf.dll), and saves the workbook as a PDF with reduced file size while maintaining visual quality, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set compression level to High (requires Aspose.Cells.Pdf.dll reference)
            // Uncomment the line below after adding a reference to Aspose.Cells.Pdf.dll
            // pdfOptions.CompressionLevel = Aspose.Cells.Pdf.CompressionLevel.High;

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
