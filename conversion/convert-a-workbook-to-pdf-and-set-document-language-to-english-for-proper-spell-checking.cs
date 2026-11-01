// Title: Convert Excel workbook to PDF and set document language to English with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set workbook built-in document property language before PDF conversion | Convert XLSX to PDF with English language setting for spell checking using Aspose.Cells | Aspose.Cells PDF conversion preserve spell‑checking language property in .NET | Set document language to English in Aspose.Cells workbook prior to saving as PDF | Aspose.Cells PdfSaveOptions default conversion example with language property
// Tags: Aspose.Cells set built-in document property language | Excel to PDF conversion Aspose.Cells C# | PdfSaveOptions default usage Aspose.Cells | Workbook language property English Aspose.Cells | Spell checking language setting PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input XLSX file, loads it into an Aspose.Cells Workbook, assigns the built‑in "Language" property to "English" when available, and saves the workbook as a PDF using default PdfSaveOptions, with error handling for runtime exceptions.
class WorkbookToPdf
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the document language to English for proper spell checking
            // Use the property name indexer as the enum may not be available in all versions
            if (workbook.BuiltInDocumentProperties["Language"] != null)
            {
                workbook.BuiltInDocumentProperties["Language"].Value = "English";
            }

            // Prepare PDF save options (default options are sufficient for conversion)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
