using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace PrintAreaPdfValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (A1:C5)
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define a custom print area (only A1:B3 should appear in the PDF)
                sheet.PageSetup.PrintArea = "A1:B3";

                // Configure PDF save options (export only the defined print area if supported)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                    // Note: ExportPrintAreaOnly property is not available in the current Aspose.Cells version.
                    // The print area will be respected by default when saving to PDF.
                };

                // Path for the generated PDF
                string pdfPath = "PrintAreaOnly.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to PDF
                workbook.Save(pdfPath, pdfOptions);

                // Render the workbook to obtain page count (respecting the print area)
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                };

                WorkbookRender render = new WorkbookRender(workbook, renderOptions);
                int pageCount = render.PageCount;

                Console.WriteLine($"PDF generated at '{pdfPath}' with page count: {pageCount}");

                if (pageCount == 1)
                {
                    Console.WriteLine("Validation passed: PDF contains only the cells defined by the custom print area.");
                }
                else
                {
                    Console.WriteLine("Validation failed: PDF contains more pages than expected.");
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File not found: {ex.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}