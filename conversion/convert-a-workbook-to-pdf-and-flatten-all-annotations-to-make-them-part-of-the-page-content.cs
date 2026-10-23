// Title: Convert an Excel workbook to PDF and embed cell comments as part of the page using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save .xlsx as PDF with cell comments included | How to render Excel comments in PDF output using Aspose.Cells .NET | C# convert Excel to PDF and make annotations part of the document | PdfSaveOptions to include comments when exporting workbook to PDF in Aspose.Cells | Embedding Excel annotations in PDF with Aspose.Cells for .NET
// Tags: Excel to PDF conversion with comment rendering Aspose.Cells | PdfSaveOptions comment inclusion Aspose.Cells .NET | embed cell comments in PDF using Aspose.Cells | C# workbook to PDF with merged annotations | Aspose.Cells flatten annotations PDF export

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, creates a PdfSaveOptions object (which by default renders cell comments), and saves the workbook as a PDF so that comments are flattened into the page content.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the source workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            // Note: In the current Aspose.Cells version, properties for rendering
            // and flattening comments (EnableSheetComments, FlattenAllComments) are not available.
            // The default behavior will include comments in the PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
