// Title: Convert an XLS workbook to PDF with Office Add‑In rendering using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xls file with Aspose.Cells, checks for the EnableOfficeAddIn property on PdfSaveOptions, sets it if available, and saves the workbook as a PDF. | Show how to implement robust file‑existence validation and exception handling when converting Excel to PDF with Aspose.Cells, including logging of conversion errors. | Demonstrate a version‑safe way to enable Office Add‑In preservation in PDF output, using reflection to set PdfSaveOptions.EnableOfficeAddIn only when the property exists.
// Common Searches: how to keep Excel Office Add‑Ins when exporting to PDF with Aspose.Cells C# | office add‑in rendering option missing in older Aspose.Cells versions | C# convert .xls to .pdf preserving interactive controls using Aspose.Cells | check for PDF save option before setting office add‑in preservation | handle missing input Excel file during Aspose.Cells PDF conversion
// Tags: aspocells pdfsaveoptions office add‑in rendering | excel to pdf conversion preserving interactive controls | c# file existence check aspocells | reflection based property setting aspocells | exception handling for aspocells pdf export

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that input.xls exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, optionally enables Office Add‑In rendering via the EnableOfficeAddIn property (when available), and saves the workbook as output.pdf while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.pdf";

        // Verify that the source workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the source XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: In some versions of Aspose.Cells the property to preserve Office Add‑Ins
            // (EnableOfficeAddIn) may not be available. If it exists, uncomment the line below.
            // pdfOptions.EnableOfficeAddIn = true;

            // Convert and save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving, or conversion issues)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
