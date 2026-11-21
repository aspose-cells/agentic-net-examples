// Title: Set worksheet page orientation to landscape before exporting a chart to PDF using Aspose.Cells for .NET
// Common Searches: how to export an Excel chart to PDF in landscape mode with Aspose.Cells C# | Aspose.Cells set worksheet orientation to landscape for PDF conversion | C# render chart full width on PDF page using Aspose.Cells | PdfSaveOptions OnePagePerSheet landscape orientation Aspose.Cells example | convert chart from .xlsx to landscape PDF using Aspose.Cells .NET
// Tags: worksheet page orientation landscape Aspose.Cells | PdfSaveOptions OnePagePerSheet chart export | full width chart PDF export Aspose.Cells | C# Aspose.Cells PDF conversion configuration | set worksheet orientation before PDF rendering

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, changes the first worksheet's page orientation to landscape, configures PdfSaveOptions to place the entire sheet on a single page, and saves the workbook—including its chart—as a PDF, ensuring the chart uses the full page width.
class Program
{
    static void Main()
    {
        // Load the workbook that contains the chart
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // OPTIONAL: Access the chart if you need to manipulate it further
        // Chart chart = sheet.Charts[0];

        // Set the page orientation of the worksheet to Landscape.
        // This orientation will be used when the worksheet is rendered to PDF,
        // ensuring the chart can use the full width of the page.
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Configure PDF save options.
        // OnePagePerSheet forces the entire sheet (including the chart) onto a single page,
        // which works well with Landscape orientation.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true
        };

        // Save the workbook (including the chart) as a PDF file.
        workbook.Save("output.pdf", pdfOptions);
    }
}
