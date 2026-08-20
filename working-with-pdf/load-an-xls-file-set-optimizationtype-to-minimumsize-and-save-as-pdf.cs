// Title: C# – Convert XLS to PDF with MinimumSize optimization using Aspose.Cells
// Description: Load an XLS workbook, set PdfSaveOptions.OptimizationType to MinimumSize, and save the file as a compact PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells XLS to PDF | PdfSaveOptions MinimumSize | C# PDF optimization Aspose | reduce PDF size Excel conversion | .NET Aspose.Cells PDF output
// Common Searches: Aspose.Cells set PDF optimization to MinimumSize | convert .xls to .pdf with smallest file size | C# Aspose.Cells PDF save options example | how to shrink PDF generated from Excel using Aspose
// Developer Intent: The developer needs to read an existing XLS file, apply the MinimumSize PDF optimization, and generate a PDF document.
// Use Cases: Create lightweight PDF reports from legacy Excel templates for email attachments. | Publish Excel‑derived PDFs on websites where bandwidth is limited. | Batch‑process a folder of XLS files into minimized PDFs on a server.
// AI Prompts: Write C# code that converts an XLS workbook to PDF with MinimumSize optimization and includes try‑catch error handling. | Explain the impact of PdfOptimizationType.MinimumSize on PDF quality, file size, and rendering speed compared to other optimization types. | Provide a script that scans a directory for .xls files and converts each to a MinimumSize PDF using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an XLS workbook, set PdfSaveOptions.OptimizationType to MinimumSize, and save the file as a compact PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing XLS file
        Workbook workbook = new Workbook("input.xls");

        // Create PDF save options and set the optimization type to MinimumSize
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
