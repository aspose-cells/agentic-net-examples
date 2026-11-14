// Title: Export a merged Excel workbook containing charts and images to PDF with Aspose.Cells for .NET
// Common Searches: how to save a combined Excel workbook as PDF with charts using Aspose.Cells C# | Aspose.Cells export workbook with embedded images to PDF preserving layout | C# set PdfSaveOptions AllColumnsInOnePagePerSheet when converting Excel to PDF | export merged workbook to PDF with chart fidelity Aspose.Cells .NET | convert Excel file that contains charts and pictures to PDF using Aspose.Cells
// Tags: Aspose.Cells export workbook to PDF with charts | PdfSaveOptions AllColumnsInOnePagePerSheet C# | merged Excel workbook PDF conversion Aspose.Cells | preserve chart and image fidelity Aspose.Cells PDF | convert combined workbook to PDF .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorkbookToPdf
{
    // Loads a merged workbook (CombinedWorkbook.xlsx) that includes charts and images, configures PdfSaveOptions to fit all columns on one page per sheet, and saves it as CombinedWorkbook.pdf using Aspose.Cells for .NET, with error handling for missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the combined workbook that contains charts and images
            string workbookPath = "CombinedWorkbook.xlsx";

            // Verify that the workbook file exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at path '{workbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Configure PDF save options (e.g., fit all columns on one page per sheet)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    AllColumnsInOnePagePerSheet = true
                };

                // Export the workbook to PDF
                string pdfPath = "CombinedWorkbook.pdf";
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Workbook exported to PDF successfully: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
