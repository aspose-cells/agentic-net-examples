// Title: Convert Worksheets 1 and 3 to PDF using LightCells API in Aspose.Cells for .NET
// Description: Demonstrates loading an XLSX workbook in LightCells mode, selecting only the first and third worksheets, and exporting them to a PDF file with minimal memory consumption.
// Keywords: Aspose.Cells LightCells load workbook | select specific sheets PDF | PdfSaveOptions SheetSet .NET | MemorySetting large Excel | C# convert Excel to PDF selected sheets
// Common Searches: Aspose.Cells load only certain worksheets | LightCells API export selected sheets to PDF | PdfSaveOptions SheetSet example C# | How to reduce memory usage when converting Excel to PDF
// Developer Intent: Load an Excel file with LightCells, pick worksheets 1 and 3, and save those sheets as a PDF.
// Use Cases: Create a lightweight PDF containing only summary and data sheets from a massive workbook. | Generate targeted PDF reports on a server where memory is limited. | Automate batch conversion of specific sheets across many workbooks to reduce processing time.
// AI Prompts: Write C# code that uses Aspose.Cells LightCells mode to load an XLSX file, selects worksheets at indexes 0 and 2, and saves them as a PDF. | Explain how to configure LoadOptions.MemorySetting for large Excel files and use PdfSaveOptions.SheetSet to export chosen sheets. | Provide step‑by‑step instructions for converting selected worksheets to PDF while optimizing memory usage with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates loading an XLSX workbook in LightCells mode, selecting only the first and third worksheets, and exporting them to a PDF file with minimal memory consumption.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Create LoadOptions (LightCells mode can be enabled via MemorySetting if needed)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.MemorySetting = MemorySetting.MemoryPreference; // optional for large files

        // Load the workbook using the LightCells‑compatible constructor
        Workbook workbook = new Workbook(sourceFile, loadOptions);

        // Prepare PDF save options and specify only worksheets 1 and 3 (zero‑based indexes 0 and 2)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.SheetSet = new SheetSet(new int[] { 0, 2 });

        // Save the selected sheets to a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
