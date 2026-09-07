// Title: Export an Excel workbook containing hidden rows to PDF with gridlines using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets PageSetup.PrintGridlines = true for every worksheet, and saves the workbook as a PDF while ensuring hidden rows do not affect the displayed gridlines. | Create a reusable C# method named ExportToPdfWithGridlines that accepts input and output paths, checks file existence, configures PdfSaveOptions, and performs the hidden‑row‑aware PDF conversion with Aspose.Cells. | Adapt the given Aspose.Cells example to add error handling for missing files and to explicitly enable gridline printing in the PDF output without altering the original worksheet formatting.
// Common Searches: Aspose.Cells hide rows but keep gridlines when converting to PDF | How to print gridlines for all worksheets in a .NET Excel to PDF conversion | C# PDF export from Excel ignoring hidden rows using Aspose.Cells | Enable PrintGridlines in Aspose.Cells before saving workbook as PDF | Validate input Excel file existence before Aspose.Cells PDF conversion
// Tags: Aspose.Cells PDF export with gridlines | hidden rows excluded from gridline rendering | PageSetup.PrintGridlines setting | C# PdfSaveOptions for Excel conversion | input file validation for Aspose.Cells workflow | apply settings to all worksheets

using Aspose.Cells;
using System;
using System.IO;

// The program verifies that the source XLSX file exists, loads it into an Aspose.Cells Workbook, iterates through each worksheet to enable PageSetup.PrintGridlines, configures default PdfSaveOptions, and saves the workbook as a PDF where hidden rows do not interfere with gridline rendering.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains hidden rows
            Workbook workbook = new Workbook(inputPath);

            // Ensure hidden rows do not affect gridline rendering
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PrintGridlines = true;
            }

            // Configure PDF save options (gridlines are already enabled via PageSetup)
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
