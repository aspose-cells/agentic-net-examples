// Title: C# – Convert Aspose.Cells Workbook to PDF and Understand Viewer Toolbar Limitations
// Description: Shows how to create or load an Aspose.Cells Workbook, add data, configure PdfSaveOptions, ensure the target folder exists, and save the workbook as a PDF. The sample also clarifies that the current Aspose.Cells API does not expose ViewerPreferences such as HideToolbar, and points to possible work‑arounds.
// Keywords: Aspose.Cells PDF conversion C# | Workbook to PDF Aspose.Cells | PdfSaveOptions C# | hide toolbar PDF Aspose | viewer preferences PDF Aspose.Cells | export Excel to PDF .NET | Aspose.Cells API limitation viewer preferences | C# create output directory | Aspose.Cells save PDF | PDF viewer UI control
// Common Searches: Aspose.Cells hide toolbar when exporting to PDF | PDF viewer preferences not supported in Aspose.Cells | C# convert Excel workbook to PDF with Aspose.Cells | How to set PdfSaveOptions in Aspose.Cells | Create output folder before saving PDF Aspose.Cells
// Developer Intent: Export an Excel workbook to PDF using Aspose.Cells for .NET and learn why the viewer toolbar cannot be hidden directly through the API.
// Use Cases: Generate a PDF report from a dynamically built workbook while managing file‑system paths. | Batch‑process multiple workbooks into PDFs, automatically creating missing output directories. | Understand API constraints around PDF viewer UI settings and plan post‑processing if toolbar hiding is required.
// AI Prompts: Write C# code with Aspose.Cells that converts a workbook to PDF and explains the lack of HideToolbar support. | Suggest a post‑processing approach (e.g., using iTextSharp or PdfSharp) to hide the toolbar in PDFs created by Aspose.Cells. | Provide a concise example that creates a workbook, adds data, ensures the output directory exists, and saves it as PDF with available PdfSaveOptions.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create or load an Aspose.Cells Workbook, add data, configure PdfSaveOptions, ensure the target folder exists, and save the workbook as a PDF. The sample also clarifies that the current Aspose.Cells API does not expose ViewerPreferences such as HideToolbar, and points to possible work‑arounds.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: ViewerPreferences (e.g., HideToolbar) are not available in the current Aspose.Cells API.
            // If needed, configure other PDF options here.

            // Define output file path
            string outputPath = "Converted.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
