// Title: Convert an Excel workbook to PDF with default RenderOfficeAddIns setting using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# convert .xlsx to .pdf with default RenderOfficeAddIns | How to include Office Add‑Ins when saving Excel as PDF using Aspose.Cells | Default PDF conversion settings in Aspose.Cells C# example | Render Office Add‑Ins in PDF output with Aspose.Cells workbook.Save | C# Aspose.Cells PDF export preserving Office Add‑Ins without custom options
// Tags: Aspose.Cells default PDF rendering options | RenderOfficeAddIns API C# | Excel workbook to PDF conversion Aspose.Cells | Preserve Office Add-Ins during PDF export | Workbook.Save PDF with default settings C#

using System;
using Aspose.Cells;

// Loads input.xlsx with Aspose.Cells, applies the default rendering configuration (including RenderOfficeAddIns), and saves the workbook as output.pdf.
class Program
{
    static void Main()
    {
        // Load an existing Excel workbook.
        // The default settings (including RenderOfficeAddIns) are used.
        Workbook workbook = new Workbook("input.xlsx");

        // Convert the workbook to PDF using the default rendering options.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
