// Title: Export an Excel workbook containing multiple pivot tables to a single PDF while preserving worksheet layout using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook with several pivot tables to one PDF | How to keep original worksheet pagination when converting Excel to PDF with Aspose.Cells | Aspose.Cells PdfSaveOptions OnePagePerSheet false example for pivot tables | Save Excel file that has pivot tables as PDF without flattening layout in .NET | Export all sheets containing pivot tables to a single PDF preserving layout using Aspose.Cells
// Tags: Aspose.Cells export workbook to PDF with pagination | PdfSaveOptions OnePagePerSheet false C# | preserve pivot table layout Aspose.Cells | multiple pivot tables single PDF output | C# convert Excel to PDF retaining sheet layout

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook that includes multiple pivot tables, configures PdfSaveOptions with OnePagePerSheet set to false to retain each worksheet's original pagination, and saves the entire workbook as a single PDF file, handling file existence checks and output directory creation.
class ExportPivotTablesToPdf
{
    static void Main()
    {
        // Path to the source Excel workbook that contains multiple pivot tables
        string sourceFile = @"C:\Input\PivotWorkbook.xlsx";

        // Path where the resulting PDF will be saved
        string pdfFile = @"C:\Output\PivotWorkbook.pdf";

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Error: Source file not found at '{sourceFile}'.");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfFile);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourceFile);

            // Configure PDF save options to preserve the original layout of each worksheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not force each worksheet onto a single page; keep original pagination
                OnePagePerSheet = false
                // The AllColumnsInOnePage property is not required for current Aspose.Cells versions
            };

            // Save the entire workbook (all worksheets) to a single PDF file
            workbook.Save(pdfFile, pdfOptions);

            Console.WriteLine("Workbook with pivot tables exported to PDF successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
