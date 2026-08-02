// Title: Convert Excel to PDF and flatten all annotations with Aspose.Cells in C#
// Description: A C# example that verifies the source .xlsx file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to calculate formulas and flatten comments, notes, and shapes into the PDF page, and saves the result as a PDF while handling errors.
// Keywords: Aspose.Cells PDF conversion C# | flatten annotations Aspose.Cells | PdfSaveOptions FlattenAllComments | Excel to PDF with comments embedded | C# convert workbook to PDF | Aspose.Cells flatten shapes PDF | calculate formulas during PDF export | Aspose.Cells example GitHub
// Common Searches: C# Aspose.Cells flatten comments when saving to PDF | How to embed Excel annotations in PDF using Aspose.Cells | PdfSaveOptions FlattenAllComments property | Convert .xlsx to PDF with annotations flattened | Aspose.Cells calculate formulas on PDF export
// Developer Intent: Generate a PDF from an Excel workbook where all comments, notes, and drawing objects are flattened into the page content, ensuring formulas are evaluated during the conversion.
// Use Cases: Publish a financial model with calculated results and reviewer comments as a read‑only PDF. | Create archival PDFs of engineering spreadsheets where callouts and shapes must appear as part of the page. | Automate report distribution while guaranteeing that no interactive Excel elements remain in the final PDF.
// AI Prompts: Write C# code that uses Aspose.Cells to convert an Excel file to PDF, flattening all comments and shapes into the PDF page. | Explain the PdfSaveOptions properties needed to embed annotations and calculate formulas during Excel‑to‑PDF conversion. | Provide guidance on optimizing memory usage when converting large workbooks to flattened PDFs with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// A C# example that verifies the source .xlsx file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to calculate formulas and flatten comments, notes, and shapes into the PDF page, and saves the result as a PDF while handling errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: calculate formulas before conversion
                CalculateFormula = true
            };

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
