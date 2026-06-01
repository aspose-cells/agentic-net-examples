using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Saving;

class Program
{
    static void Main()
    {
        // Set the current thread to Arabic culture (e.g., Saudi Arabia)
        CultureInfo arabicCulture = new CultureInfo("ar-SA");
        System.Threading.Thread.CurrentThread.CurrentCulture = arabicCulture;
        System.Threading.Thread.CurrentThread.CurrentUICulture = arabicCulture;

        // Load the existing workbook (replace with your actual file path)
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for the worksheet
        worksheet.DisplayRightToLeft = true;

        // Determine the data range (assumes data starts at A1 and occupies the used area)
        int startRow = 0;
        int startColumn = 0;
        int endRow = worksheet.Cells.MaxDataRow;
        int endColumn = worksheet.Cells.MaxDataColumn;
        CellArea dataArea = CellArea.CreateCellArea(startRow, startColumn, endRow, endColumn);

        // Add subtotals:
        //   - Group by the first column (index 0)
        //   - Use SUM function on the second column (index 1)
        //   - Replace existing subtotals, no page breaks, summary below data
        worksheet.Cells.Subtotal(
            dataArea,
            0,                                 // groupBy column index
            ConsolidationFunction.Sum,         // subtotal function
            new int[] { 1 },                   // columns to subtotal
            true,                              // replace existing subtotals
            false,                             // no page breaks between groups
            true                               // place summary below data
        );

        // Configure PDF save options (set a font that supports Arabic characters)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DefaultFont = "Arial",            // fallback font for Arabic text
            CheckWorkbookDefaultFont = true   // ensure default font is checked
        };

        // Save the workbook as PDF
        string outputFile = "output.pdf";
        workbook.Save(outputFile, pdfOptions);
    }
}