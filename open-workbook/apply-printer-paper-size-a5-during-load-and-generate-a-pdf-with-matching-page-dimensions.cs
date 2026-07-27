// Title: Aspose.Cells for .NET – Load Excel with A5 paper size and export to PDF
// Description: Demonstrates how to use Aspose.Cells LoadOptions to set the default paper size to A5 when opening an Excel workbook, optionally enforce A5 on each worksheet, and save the workbook as a PDF with exact A5 dimensions. The example is written in C# and works with .NET 6+.
// Keywords: Aspose.Cells C# | LoadOptions SetPaperSize | A5 paper size Excel | Export Excel to PDF | PDF page dimensions A5 | Worksheet PageSetup A5 | Aspose.Cells .NET example
// Common Searches: set A5 paper size when loading Excel with Aspose.Cells | export Excel to PDF A5 page size .NET | Aspose.Cells LoadOptions SetPaperSize example | how to force worksheet page size before PDF conversion | C# Aspose.Cells PDF page dimensions
// Developer Intent: Apply A5 paper size during workbook load and generate a PDF that matches those dimensions.
// Use Cases: Create printable A5 flyers or brochures directly from Excel templates. | Produce compact A5 financial statements for mobile distribution. | Batch‑convert multiple workbooks to A5‑sized PDFs for consistent publishing.
// AI Prompts: Write C# code using Aspose.Cells to load an Excel file with A5 paper size and save it as a PDF. | Explain how to guarantee every worksheet retains the A5 size during PDF export with Aspose.Cells. | Adapt the example to use Letter paper size while keeping the same load‑and‑save workflow.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LoadOptions to set the default paper size to A5 when opening an Excel workbook, optionally enforce A5 on each worksheet, and save the workbook as a PDF with exact A5 dimensions. The example is written in C# and works with .NET 6+.
class Program
{
    static void Main()
    {
        // Create load options and set the default paper size to A5
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.SetPaperSize(PaperSizeType.PaperA5);

        // Load the workbook with the specified load options
        // (replace "input.xlsx" with the path to your source file)
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Ensure each worksheet uses A5 paper size (optional but guarantees consistency)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;
        }

        // Save the workbook as PDF; the PDF pages will have A5 dimensions
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
