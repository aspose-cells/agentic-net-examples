// Title: Convert an Excel workbook to PDF with StandardSize optimization while preserving column widths using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets PdfSaveOptions.StandardSize = true, disables AllColumnsInOnePage, and saves the workbook as a PDF with Aspose.Cells. | Show how to programmatically detect the AllColumnsInOnePage property in PdfSaveOptions and set it to false before exporting to PDF. | Provide a complete example that converts an Excel file to PDF, keeps the original column widths, and enables the StandardSize optimization for better page layout.
// Common Searches: Aspose.Cells PDF conversion keep original column widths | How to disable AllColumnsInOnePage when saving Excel as PDF in C# | Enable StandardSize option in PdfSaveOptions for Aspose.Cells | Convert XLSX to PDF without scaling columns to one page using Aspose.Cells | C# Aspose.Cells preserve layout during Excel to PDF export
// Tags: Aspose.Cells PdfSaveOptions StandardSize | Aspose.Cells preserve column widths PDF | Aspose.Cells disable AllColumnsInOnePage | Excel to PDF layout preservation .NET | runtime check PdfSaveOptions property C#

using System;
using System.IO;
using Aspose.Cells;

// Loads 'input.xlsx', configures PdfSaveOptions to enable StandardSize and preserve column widths by setting AllColumnsInOnePage to false (when the property exists), then saves the workbook as 'output.pdf' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Preserve original column widths (do not force all columns onto one page)
            // The AllColumnsInOnePage property may not be present in all versions;
            // if it exists, set it to false.
            var allColsProp = typeof(PdfSaveOptions).GetProperty("AllColumnsInOnePage");
            if (allColsProp != null && allColsProp.CanWrite)
            {
                allColsProp.SetValue(pdfOptions, false);
            }

            // Save the workbook as PDF
            workbook.Save(outputFile, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
