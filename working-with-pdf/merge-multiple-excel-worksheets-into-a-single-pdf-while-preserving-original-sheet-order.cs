// Title: Combine all worksheets of an Excel workbook into a single PDF while keeping the original sheet order using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, configures PdfSaveOptions to merge all worksheets, and saves the output as one PDF preserving the sheet sequence. | Show how to check for the existence of the source Excel file, create the destination folder if needed, and export the workbook to a concatenated PDF without using OnePagePerSheet. | Provide a C# example that sets PdfSaveOptions.OnePagePerSheet = false to produce a single PDF from multiple sheets in their original order.
// Common Searches: how to export multiple Excel sheets to a single PDF using Aspose.Cells C# | preserve worksheet order when converting an Excel workbook to PDF with Aspose.Cells | Aspose.Cells PdfSaveOptions merge sheets into one PDF .NET | C# convert Excel workbook with many sheets to one PDF file | save Excel workbook as concatenated PDF Aspose.Cells example
// Tags: Aspose.Cells PDF conversion with sheet concatenation | PdfSaveOptions configuration for multi-sheet PDF | C# preserve worksheet order in PDF export | Excel workbook to single PDF using Aspose.Cells | create output folder before saving PDF in C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies the input .xlsx file, ensures the output directory exists, loads the workbook with Aspose.Cells, sets PdfSaveOptions.OnePagePerSheet to false so all worksheets are concatenated in their original order, and saves the result as a single PDF file.
class ExcelToPdfMerger
{
    static void Main()
    {
        // Path to the source Excel file containing multiple worksheets
        string excelPath = @"C:\Input\WorkbookWithMultipleSheets.xlsx";

        // Path where the merged PDF will be saved
        string pdfPath = @"C:\Output\MergedWorkbook.pdf";

        try
        {
            // Verify that the input Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the Excel workbook (preserves original sheet order)
            Workbook workbook = new Workbook(excelPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Allow sheets to span multiple pages and concatenate them in order
                OnePagePerSheet = false
                // Additional options can be set here, e.g., GridlineType, Headings, etc.
            };

            // Save the workbook as a single PDF file using the options
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine("Workbook successfully merged into PDF: " + pdfPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
