// Title: Convert an Excel workbook to PDF with system default font fallback using Aspose.Cells in C#
// Common Searches: Aspose.Cells convert xlsx to pdf with system font fallback in C# | C# code to export Excel workbook to PDF using default font substitution | How to handle missing fonts when saving Excel as PDF with Aspose.Cells | Set default font for PDF conversion in Aspose.Cells C# example | Excel to PDF conversion preserving layout when fonts are unavailable Aspose.Cells
// Tags: Aspose.Cells default font substitution PDF | Excel to PDF conversion C# Aspose.Cells | Workbook.Save PDF with system font fallback | handling missing fonts Aspose.Cells conversion | SaveFormat.Pdf usage Aspose.Cells | system default font for PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, optionally relies on the operating system's default font for missing typefaces, and saves the workbook as a PDF using Aspose.Cells, ensuring proper font substitution during conversion.
class WorkbookToPdfConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Optionally, specify a default font for substitution.
            // If not set, Aspose.Cells uses the system's default font.
            // workbook.Settings.DefaultFont = "Arial";

            // Save the workbook as PDF.
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
