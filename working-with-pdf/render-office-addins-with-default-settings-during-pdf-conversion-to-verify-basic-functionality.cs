// Title: Convert an Excel workbook to PDF with default settings that automatically render Office Add‑Ins using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as a PDF with Aspose.Cells using the default PdfSaveOptions so that any Office Add‑Ins are rendered. | Show how to export an Excel workbook to PDF in C# without modifying PdfSaveOptions, relying on the built‑in behavior to include Office Add‑Ins.
// Common Searches: how to export Excel to PDF with Office Add‑Ins using Aspose.Cells C# | Aspose.Cells default PDF conversion renders add‑ins | C# convert .xlsx to PDF preserving Office Add‑Ins Aspose.Cells
// Tags: Aspose.Cells default PdfSaveOptions rendering add‑ins | C# export Excel to PDF with Office Add‑Ins | PdfSaveOptions without custom settings Aspose.Cells | Excel workbook PDF conversion preserving add‑ins .NET | Aspose.Cells PDF export default behavior

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads an Excel workbook (input.xlsx) and saves it as a PDF (output.pdf) using Aspose.Cells' PdfSaveOptions with default settings, which automatically render any Office Add‑Ins present in the workbook.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options with default settings.
        // Default behavior renders Office Add‑Ins, so no additional configuration is required.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF using the default options.
        workbook.Save("output.pdf", pdfOptions);
    }
}
