// Title: C# – Add Arabic RTL Subtotals to Excel and Export as PDF with Aspose.Cells
// Description: Loads an Excel workbook, applies the ar‑SA CultureInfo for Arabic number/date formats, enables right‑to‑left display, creates SUM subtotals grouped by the first column, configures PdfSaveOptions with an Arabic‑compatible font, and saves the result as a PDF.
// Keywords: Aspose.Cells | C# | Arabic culture | ar-SA | right-to-left | RTL | subtotal | Excel | PDF export | PdfSaveOptions | Arabic font | Excel localization | Middle East
// Common Searches: Aspose.Cells add subtotal RTL | C# export Arabic Excel to PDF | set CultureInfo ar-SA Aspose.Cells | enable right to left worksheet Aspose | Arabic PDF export Aspose.Cells | subtotal function in Aspose.Cells C#
// Developer Intent: Create a PDF from an Excel workbook that uses Arabic (ar‑SA) culture, displays right‑to‑left layout, and includes automatically calculated subtotals.
// Use Cases: Generate financial statements for Saudi Arabian or Middle‑Eastern markets with Arabic number formats, RTL subtotals, and PDF delivery. | Automate Arabic‑language invoices where items are grouped, subtotaled, and exported as PDF for printing or archiving. | Produce localized sales summaries that group data, add subtotals, and render correctly in PDF using Arabic fonts.
// AI Prompts: Write C# code with Aspose.Cells to load an Excel file, set CultureInfo to ar‑SA, enable right‑to‑left display, add SUM subtotals grouped by the first column, and save as PDF using an Arabic‑compatible font. | Explain how PdfSaveOptions should be configured in Aspose.Cells to ensure Arabic characters render correctly when exporting to PDF. | Show how to determine the data range dynamically for Subtotal based on MaxDataRow and MaxDataColumn in Aspose.Cells.

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Loads an Excel workbook, applies the ar‑SA CultureInfo for Arabic number/date formats, enables right‑to‑left display, creates SUM subtotals grouped by the first column, configures PdfSaveOptions with an Arabic‑compatible font, and saves the result as a PDF.
class ArabicRtlSubtotalPdf
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the workbook's culture to Arabic (Saudi Arabia)
        // This influences number/date formatting according to Arabic conventions
        workbook.Settings.CultureInfo = new CultureInfo("ar-SA");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for the worksheet
        worksheet.DisplayRightToLeft = true;

        // Define the data range for subtotals (assumes data starts at A1)
        // Adjust the range as needed for your specific data layout
        int startRow = 0;               // zero‑based index for row 1
        int startColumn = 0;            // zero‑based index for column A
        int endRow = worksheet.Cells.MaxDataRow;       // last row with data
        int endColumn = worksheet.Cells.MaxDataColumn; // last column with data
        CellArea dataArea = CellArea.CreateCellArea(startRow, startColumn, endRow, endColumn);

        // Create subtotals:
        // - Group by the first column (index 0)
        // - Use SUM function
        // - Apply subtotal to the second column (index 1)
        // - Replace existing subtotals, no page breaks, summary placed below data
        worksheet.Cells.Subtotal(
            dataArea,
            0,
            ConsolidationFunction.Sum,
            new int[] { 1 },
            true,   // replace existing subtotals
            false,  // no page breaks between groups
            true    // place summary below data
        );

        // Configure PDF save options for proper Arabic rendering
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use a font that supports Arabic characters
            DefaultFont = "Arial",
            // Ensure the default workbook font is checked for missing glyphs
            CheckWorkbookDefaultFont = true
        };

        // Save the workbook as PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
