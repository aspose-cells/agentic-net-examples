using System;
using Aspose.Cells;
using Aspose.Cells.Saving;   // PdfSaveOptions resides here

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells M5:N6
        // M5 -> row 4, column 12 (zero‑based)
        // N6 -> row 5, column 13 (zero‑based)
        // totalRows = 2, totalColumns = 2
        cells.Merge(4, 12, 2, 2);

        // Apply Accounting number format (built‑in format 44) to the merged cell
        Style accountingStyle = workbook.CreateStyle();
        accountingStyle.Number = 44;               // Accounting format
        cells[4, 12].SetStyle(accountingStyle);    // Upper‑left cell of the merged area

        // Configure PDF save options (example: export document structure)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        // Set page margins (in inches) via the worksheet's PageSetup
        worksheet.PageSetup.LeftMargin = 0.5;
        worksheet.PageSetup.RightMargin = 0.5;
        worksheet.PageSetup.TopMargin = 0.5;
        worksheet.PageSetup.BottomMargin = 0.5;

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}