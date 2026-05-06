using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class FontSubstitutionAnalysis
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Desired output PDF file path
        string pdfPath = "output.pdf";

        // Load options – specify the format of the source file (optional but explicit)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Configure PDF save options to capture font‑substitution details
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Enable per‑character font substitution analysis
        pdfOptions.IsFontSubstitutionCharGranularity = true;

        // Try to use the workbook's default font first when a character's font is missing
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Set a fallback default font (e.g., Arial) for characters without a suitable font
        pdfOptions.DefaultFont = "Arial";

        // Ensure the engine checks font compatibility for every character
        pdfOptions.CheckFontCompatibility = true;

        // Perform the conversion using the overload that accepts load and save options
        ConversionUtility.Convert(sourcePath, loadOptions, pdfPath, pdfOptions);

        Console.WriteLine("Excel file has been converted to PDF with font‑substitution analysis enabled.");
    }
}