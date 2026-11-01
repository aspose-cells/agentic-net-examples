// Title: Export an Excel workbook with multiple pivot tables to a single PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook containing several pivot tables to one PDF | How to save an Excel file with multiple pivot tables as a single PDF in .NET | C# PdfSaveOptions settings for preserving pivot table layout during PDF conversion | Export all worksheets with pivot tables to a single PDF using Aspose.Cells | Avoid page breaks for pivot tables when converting Excel to PDF with Aspose.Cells C#
// Tags: Aspose.Cells export workbook to PDF | PdfSaveOptions preserve pivot layout | C# convert Excel pivot tables to single PDF | OnePagePerSheet false Aspose.Cells | AllColumnsInOnePagePerSheet false PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing Excel workbook that contains multiple pivot tables, configures PdfSaveOptions to keep the original layout (disabling OnePagePerSheet and AllColumnsInOnePagePerSheet), and saves the entire workbook—including all worksheets and their pivot tables—to a single PDF file, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains multiple pivot tables.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not force each worksheet onto a single page.
                OnePagePerSheet = false,
                // Keep column widths as they appear; do not shrink them to fit one page.
                AllColumnsInOnePagePerSheet = false
                // Note: RenderPivotTable property is not available in this version of Aspose.Cells.
            };

            // Export the entire workbook (all sheets with their pivot tables) to a single PDF file.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
