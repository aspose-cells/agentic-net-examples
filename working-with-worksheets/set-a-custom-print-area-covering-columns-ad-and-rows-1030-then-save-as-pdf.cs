// Title: Set Print Area A10:D30 and Export Worksheet to PDF with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills cells in rows 10‑30 and columns A‑D, defines the print area with PageSetup.PrintArea = "A10:D30", and saves the sheet as a PDF so only that range appears in the output.
// Keywords: Aspose.Cells C# print area | custom print range PDF | export Excel to PDF Aspose | .NET workbook PDF export | set PageSetup.PrintArea | A10:D30 PDF output | Aspose.Cells PDF conversion | C# Excel print area
// Common Searches: Aspose.Cells set print area before PDF export | C# export selected range to PDF using Aspose.Cells | how to limit PDF output to A10:D30 in Aspose.Cells | save Excel worksheet as PDF with custom print area .NET | Aspose.Cells PDF only specific cells
// Developer Intent: Define a print area covering A10:D30 and generate a PDF of that range using Aspose.Cells for .NET.
// Use Cases: Produce a PDF report that contains only a summary table located in rows 10‑30, columns A‑D. | Create printable invoices or receipts by restricting the PDF to a predefined cell block. | Distribute a chart or data snapshot without exporting the entire worksheet.
// AI Prompts: Show C# code with Aspose.Cells that sets a custom print area and saves the worksheet as a PDF. | How can I export multiple print areas from one worksheet to separate PDF files using Aspose.Cells? | Explain how to adjust margins and orientation while using a custom print area for PDF conversion in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaPdfDemo
{
    // Creates a workbook, fills cells in rows 10‑30 and columns A‑D, defines the print area with PageSetup.PrintArea = "A10:D30", and saves the sheet as a PDF so only that range appears in the output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Populate some data within the print area for demonstration
            for (int row = 9; row <= 29; row++)          // Rows 10‑30 (0‑based index)
            {
                for (int col = 0; col <= 3; col++)       // Columns A‑D
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Set the custom print area to columns A‑D and rows 10‑30
            worksheet.PageSetup.PrintArea = "A10:D30";

            // Save the workbook as PDF; the defined print area will be used
            workbook.Save("CustomPrintArea.pdf", SaveFormat.Pdf);
        }
    }
}
