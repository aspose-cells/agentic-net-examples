using System;
using Aspose.Cells; // Aspose.Cells namespace provides Workbook and PdfSaveOptions

// Author: Aspose.Cells .NET example – exporting a workbook (with charts and images) to PDF
class Program
{
    static void Main()
    {
        // Load the combined workbook that contains charts and images.
        // Replace the file path with the actual location of your workbook.
        Workbook workbook = new Workbook("combined.xlsx");

        // Configure PDF save options if specific settings are required.
        // Here we use default options, which preserve charts and images.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example of a custom setting (optional):
            // OnePagePerSheet = false
        };

        // Save the workbook as a PDF file.
        // The resulting PDF can be inspected to verify visual fidelity.
        workbook.Save("combined_output.pdf", pdfOptions);

        Console.WriteLine("Workbook exported to PDF successfully.");
    }
}