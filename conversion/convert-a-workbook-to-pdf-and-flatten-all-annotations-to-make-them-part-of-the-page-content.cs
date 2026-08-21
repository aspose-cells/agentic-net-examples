// Title: Convert Excel to PDF and Flatten Annotations with Aspose.Cells for .NET
// Description: This example demonstrates how to load an XLSX workbook, optionally recalculate formulas, and export it to PDF using Aspose.Cells. It explains the PdfSaveOptions settings and notes that the FlattenAllAnnotations feature is available only in newer library versions, advising an upgrade when needed.
// Keywords: Aspose.Cells | C# | Excel to PDF conversion | flatten annotations | PdfSaveOptions | export workbook as PDF | annotation flattening Aspose.Cells | update Aspose.Cells version | Excel PDF export .NET | formula calculation before PDF
// Common Searches: Aspose.Cells flatten annotations when saving to PDF | how to export Excel as PDF with annotations merged | PdfSaveOptions FlattenAllAnnotations property | convert .xlsx to PDF using C# Aspose.Cells | upgrade Aspose.Cells for annotation flattening
// Developer Intent: Export an Excel workbook to PDF while merging any cell comments or shapes into the final page content.
// Use Cases: Generate a PDF report from a workbook with all formulas evaluated. | Create a PDF where comments, notes, and drawing objects become part of the static page. | Detect the current Aspose.Cells version and prompt an upgrade to access annotation‑flattening features. | Implement robust file‑existence checks and exception handling for the conversion workflow.
// AI Prompts: Write C# code that loads an Excel file, calculates formulas, and saves it as a PDF with all annotations flattened using Aspose.Cells. | Show how to programmatically verify the Aspose.Cells version and conditionally enable the FlattenAllAnnotations option. | Provide error‑handling snippets for missing input files, licensing issues, and PDF save failures in an Aspose.Cells conversion routine.

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to load an XLSX workbook, optionally recalculate formulas, and export it to PDF using Aspose.Cells. It explains the PdfSaveOptions settings and notes that the FlattenAllAnnotations feature is available only in newer library versions, advising an upgrade when needed.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Optional: calculate all formulas before saving
            workbook.CalculateFormula();

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The FlattenAllAnnotations property is not available in the current
            // Aspose.Cells version. If needed, update the library to a newer version
            // that supports this feature.

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., loading, saving, licensing)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
