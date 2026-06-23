using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSM workbook from disk
        Workbook workbook = new Workbook("input.xlsm");

        // Ensure that gridlines are not printed for any worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.PageSetup.PrintGridlines = false;
        }

        // Create PDF save options and set a non‑solid gridline type (optional)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            GridlineType = GridlineType.Hair // Hair is a thin line; solid lines are avoided
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}