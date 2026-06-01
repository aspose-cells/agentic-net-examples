using System;
using Aspose.Cells;

public class WorkbookToPdfConverter
{
    // Converts an Excel workbook to PDF while preserving layout and formatting.
    public static void Convert(string sourcePath, string destinationPath)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(sourcePath);

        // Recalculate all formulas to ensure up‑to‑date values.
        workbook.CalculateFormula();

        // Configure PDF save options to retain document structure and default fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,   // Preserve the document hierarchy.
            CheckWorkbookDefaultFont = true, // Use workbook default font for Unicode characters.
            OnePagePerSheet = true           // Keep each sheet on a single PDF page.
        };

        // Save the workbook as a PDF using the configured options.
        workbook.Save(destinationPath, pdfOptions);
    }

    // Example usage.
    public static void Main()
    {
        string sourceFile = "input.xlsx";
        string pdfFile = "output.pdf";

        Convert(sourceFile, pdfFile);

        Console.WriteLine("Workbook successfully converted to PDF.");
    }
}