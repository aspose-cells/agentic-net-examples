// Title: C# – Apply 95% Zoom to Worksheets with >500 Rows and Export to PDF using Aspose.Cells
// Description: Load an Excel workbook, loop through each worksheet, check the used row count with MaxDataRow, set the Zoom property to 95% for sheets that exceed 500 rows, and save the entire workbook as a PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | worksheet zoom | MaxDataRow | conditional zoom | export to PDF | .NET Excel to PDF | large worksheet scaling | batch PDF conversion | Excel row count zoom
// Common Searches: Aspose.Cells set zoom for worksheets over 500 rows | C# export Excel to PDF after adjusting zoom | How to apply conditional zoom in Aspose.Cells | Iterate worksheets and change zoom before PDF conversion | Zoom property Aspose.Cells example
// Developer Intent: Automatically set a 95% zoom on any worksheet that contains more than 500 rows and then generate a PDF of the workbook.
// Use Cases: Create printable PDFs where dense sheets are scaled to fit more content per page. | Batch‑process workbooks to normalize zoom levels for large worksheets before reporting. | Generate consistent PDF documentation from Excel files with mixed sheet sizes.
// AI Prompts: Generate C# code with Aspose.Cells that sets a 95% zoom on worksheets having over 500 rows and saves the workbook as a PDF. | Show how to iterate all worksheets, use MaxDataRow to determine row count, apply conditional zoom, and export to PDF with custom options. | Explain the steps to modify the Zoom property for selected sheets before PDF conversion using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an Excel workbook, loop through each worksheet, check the used row count with MaxDataRow, set the Zoom property to 95% for sheets that exceed 500 rows, and save the entire workbook as a PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the number of rows that contain data (zero‑based index + 1)
            int usedRows = sheet.Cells.MaxDataRow + 1;

            // If the worksheet has more than 500 rows, set the zoom to 95%
            if (usedRows > 500)
            {
                sheet.Zoom = 95; // Worksheet.Zoom property (percentage)
            }
        }

        // Create PDF save options (default settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
