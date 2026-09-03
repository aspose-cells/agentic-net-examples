// Title: How to freeze the first row and column in an Aspose.Cells worksheet and export it to PDF while preserving the frozen view (C#)
// AI Prompts: Freeze the top row and left column of a worksheet, then save the workbook as a PDF using Aspose.Cells for .NET. | Generate a PDF from an Aspose.Cells workbook that maintains the frozen pane layout, configuring PdfSaveOptions in C#. | Adjust the code to force each worksheet onto a single PDF page while keeping the frozen panes intact.
// Common Searches: Aspose.Cells C# export to PDF keep frozen panes visible | How to preserve freeze panes when converting Excel to PDF with Aspose.Cells | C# sample to freeze first row and column then save as PDF using Aspose.Cells | PdfSaveOptions OnePagePerSheet with frozen panes Aspose.Cells example | Freeze panes before PDF conversion Aspose.Cells .NET tutorial
// Tags: freeze panes pdf export aspose.cells | aspose.cells pdfsaveoptions frozen view | c# freeze first row column aspose.cells | export worksheet to pdf with frozen panes | onepagepersheet option aspose.cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with sample data, freezes the first row and column, and saves the workbook as a PDF using PdfSaveOptions, ensuring the frozen pane view is retained in the output.
class FreezePaneToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first row and first column (A1 will be the top‑left visible cell)
            sheet.FreezePanes(1, 1, 1, 1);

            // Set PDF save options (no special flags needed for frozen panes)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: keep each sheet on a single page
                // OnePagePerSheet = true
            };

            // Save the workbook as PDF while preserving the frozen view
            workbook.Save("FrozenPaneOutput.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
