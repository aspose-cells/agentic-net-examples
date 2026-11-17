// Title: C# example: Convert an Excel workbook with WordArt to PDF using Aspose.Cells and prepare for ICC‑profile‑aware gradient fills
// Common Searches: c# aspnet convert xlsx containing wordart to pdf preserving gradient fill colors | aspocells pdfsaveoptions embed icc profile for color management during pdf export | how to export an Excel workbook to pdf with color‑space awareness using Aspose.Cells | sample code for converting Excel to PDF while handling wordart objects in .net | aspocells pdf conversion gradient fill icc profile placeholder example
// Tags: Aspose.Cells PDF conversion with WordArt | PdfSaveOptions color management support | Excel to PDF gradient fill handling | ICC profile integration Aspose.Cells | C# workbook to PDF export example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file that contains WordArt, creates a PdfSaveOptions object (which currently lacks an IccProfile property), and saves the workbook as a PDF. It includes checks for the input workbook and an ICC profile file, and shows where an ICC profile could be assigned in future Aspose.Cells releases to enable color‑space‑aware gradient fills.
class SpreadsheetToPdfWithIcc
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string iccFile = "sRGB.icc";
        const string outputFile = "output.pdf";

        try
        {
            // Verify that the required files exist
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input workbook not found: {inputFile}");

            if (!File.Exists(iccFile))
                throw new FileNotFoundException($"ICC profile not found: {iccFile}");

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputFile);

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Aspose.Cells PdfSaveOptions does not expose an IccProfile property.
            // If future versions add support, assign the ICC profile here:
            // pdfOptions.IccProfile = new IccProfile(iccFile);

            // Save the workbook directly to PDF using the configured options
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
