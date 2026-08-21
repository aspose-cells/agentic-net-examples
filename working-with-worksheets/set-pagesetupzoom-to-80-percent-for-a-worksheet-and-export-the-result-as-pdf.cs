// Title: Set Worksheet Zoom to 80% and Export as PDF using Aspose.Cells for .NET
// Description: Creates a workbook, sets the first worksheet's PageSetup.Zoom to 80 % with percent‑based scaling, and saves the workbook as a PDF. The resulting PDF reflects the specified zoom level.
// Keywords: Aspose.Cells | C# worksheet zoom | PageSetup.Zoom | IsPercentScale | PDF export | custom page scaling | Aspose.Cells PDF options | .NET spreadsheet to PDF
// Common Searches: Aspose.Cells set worksheet zoom 80 percent | export worksheet to PDF with custom scaling .NET | PageSetup.IsPercentScale true Aspose.Cells | C# change page zoom before PDF conversion | how to adjust worksheet scaling for PDF output
// Developer Intent: Apply an 80 % page zoom to a worksheet and generate a PDF file.
// Use Cases: Print reports where more rows fit on each PDF page by scaling to 80 %. | Create brand‑consistent PDFs with a custom zoom while leaving other sheets at default scaling. | Improve readability of a summary sheet by reducing its zoom before exporting the workbook to PDF.
// AI Prompts: Generate C# code that sets PageSetup.Zoom to 80% and saves the workbook as a PDF with Aspose.Cells. | Explain how PageSetup.IsPercentScale influences PDF rendering in Aspose.Cells for .NET. | Show an example that exports multiple worksheets, each with a different zoom level, to separate PDF files.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Creates a workbook, sets the first worksheet's PageSetup.Zoom to 80 % with percent‑based scaling, and saves the workbook as a PDF. The resulting PDF reflects the specified zoom level.
    public class SetZoomAndExportPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default contains one worksheet)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the page scaling (zoom) to 80%
                worksheet.PageSetup.Zoom = 80;
                // Ensure the scaling mode is percent‑based
                worksheet.PageSetup.IsPercentScale = true;

                // Prepare PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF; the zoom setting will be applied to the output
                string outputPath = "Worksheet_Zoom80.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook saved as PDF with 80% zoom: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
