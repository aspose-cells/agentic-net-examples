// Title: Create a Landscape PDF from a Wide Excel Worksheet with Aspose.Cells for .NET
// Description: This example builds a workbook with 50 columns, sets the worksheet orientation to landscape, configures the page setup to fit all columns on a single page width, and saves the sheet as a PDF using PdfSaveOptions (OnePagePerSheet and AllColumnsInOnePagePerSheet). The result is a one‑page landscape PDF that preserves the full width of a wide Excel sheet.
// Keywords: Aspose.Cells landscape PDF | fit all columns one page PDF | Excel to PDF landscape orientation .NET | PdfSaveOptions OnePagePerSheet | export wide worksheet to PDF | C# Aspose.Cells PDF export
// Common Searches: Aspose.Cells export wide Excel sheet to single landscape PDF | fit all columns on one PDF page Aspose.Cells | set worksheet orientation landscape PDF Aspose | C# generate landscape PDF from Excel with Aspose | one page per sheet PDF options Aspose.Cells
// Developer Intent: Generate a landscape PDF that places every column of a wide worksheet on a single page.
// Use Cases: Produce printable financial statements with dozens of columns on a compact landscape PDF. | Convert dashboard or analytics spreadsheets into a single‑page PDF for easy email distribution. | Automate batch conversion of multiple wide worksheets into one‑page landscape PDFs for archiving.
// AI Prompts: Show how to add a custom header and footer to the landscape PDF while keeping all columns on one page. | Explain how to adjust margins and scaling so the wide sheet fits comfortably within the landscape PDF. | Provide code to export each worksheet of a multi‑sheet workbook as separate landscape PDFs with the same column‑fit settings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfLandscapeDemo
{
    // This example builds a workbook with 50 columns, sets the worksheet orientation to landscape, configures the page setup to fit all columns on a single page width, and saves the sheet as a PDF using PdfSaveOptions (OnePagePerSheet and AllColumnsInOnePagePerSheet). The result is a one‑page landscape PDF that preserves the full width of a wide Excel sheet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with wide data (many columns)
                for (int col = 0; col < 50; col++)
                {
                    // Header
                    sheet.Cells[0, col].PutValue($"Header {col + 1}");

                    // Sample rows
                    for (int row = 1; row <= 20; row++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
                    }
                }

                // Configure page setup for landscape orientation
                sheet.PageSetup.Orientation = PageOrientationType.Landscape;

                // Fit all columns to a single page width; height will adjust automatically
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 0; // 0 means auto

                // Create PDF save options that force one page per sheet
                // and place all columns on that single page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,
                    AllColumnsInOnePagePerSheet = true
                };

                string outputPath = "WideSheet_Landscape.pdf";

                // Save the workbook as a landscape PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF generated successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
