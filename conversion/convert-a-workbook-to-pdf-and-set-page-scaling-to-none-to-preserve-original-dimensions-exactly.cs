// Title: C# – Convert Excel Workbook to PDF without Scaling using Aspose.Cells
// Description: Load an .xlsx file with Aspose.Cells, set the worksheet PageSetup to percent scaling with Zoom = 100 % (no scaling), configure PdfSaveOptions (OnePagePerSheet = false), and save as a PDF that retains the exact original dimensions.
// Keywords: Aspose.Cells | C# | .NET | Excel to PDF conversion | disable page scaling | PageSetup Zoom 100 | PdfSaveOptions | preserve layout | OnePagePerSheet false
// Common Searches: Aspose.Cells export Excel to PDF without scaling | C# set page zoom 100% for PDF conversion | how to keep original size when converting workbook to PDF | PdfSaveOptions OnePagePerSheet false example | convert .xlsx to PDF preserving layout Aspose
// Developer Intent: Generate a PDF from an Excel workbook that matches the worksheet’s original size, with no automatic scaling applied.
// Use Cases: Produce printable PDFs of financial statements where the layout must be identical to the Excel source. | Batch‑convert marketing dashboards to PDF while retaining exact column widths and graphics. | Export engineering calculation sheets to PDF for documentation, ensuring dimensions remain unchanged.
// AI Prompts: Write C# code using Aspose.Cells to convert an .xlsx file to PDF with no scaling (IsPercentScale = true, Zoom = 100). | Create a reusable C# method that accepts input and output paths, disables page scaling, sets OnePagePerSheet to false, and saves the workbook as a PDF preserving original dimensions.

using System;
using Aspose.Cells;

// Load an .xlsx file with Aspose.Cells, set the worksheet PageSetup to percent scaling with Zoom = 100 % (no scaling), configure PdfSaveOptions (OnePagePerSheet = false), and save as a PDF that retains the exact original dimensions.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet's page setup
        PageSetup pageSetup = workbook.Worksheets[0].PageSetup;

        // Disable automatic scaling and set zoom to 100% to keep original dimensions
        pageSetup.IsPercentScale = true;   // Use percent scaling mode
        pageSetup.Zoom = 100;              // 100% zoom (no scaling)

        // Create PDF save options (optional, ensures default behavior)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = false; // Do not force one page per sheet

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
