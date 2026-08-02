// Title: C# – Unhide All Rows in First Worksheet (Default Height) and Export to PDF with Aspose.Cells
// Description: Load an Excel file, unhide every row in the first sheet while preserving the default row height, and convert the workbook to a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# unhide rows | export worksheet to PDF | default row height Aspose.Cells | convert Excel to PDF .NET | unhide hidden rows programmatically
// Common Searches: Aspose.Cells how to unhide all rows before PDF conversion | C# code to show hidden rows and save Excel as PDF | unhide rows default height Aspose.Cells example | convert hidden‑row Excel file to PDF with Aspose
// Developer Intent: Reveal every row in the first worksheet using the original height and generate a PDF from the workbook.
// Use Cases: Produce printable PDFs from templates where hidden rows must appear. | Batch‑process Excel reports, ensuring no data is omitted due to hidden rows. | Create compliance‑ready PDFs that display all rows regardless of prior visibility settings.
// AI Prompts: Generate C# code that opens an Excel workbook, unhides all rows in the first sheet with default height, and saves it as a PDF using Aspose.Cells. | Show how to calculate the last data row, unhide rows from index 0 to that row, and export the workbook to PDF with Aspose.Cells in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an Excel file, unhide every row in the first sheet while preserving the default row height, and convert the workbook to a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the number of rows to unhide.
        // MaxDataRow returns the last row index that contains data.
        // Adding 1 gives the total count of rows that may be hidden.
        int totalRows = cells.MaxDataRow + 1;

        // Unhide all rows starting from row 0.
        // Height = -1 keeps the original (default) row height.
        cells.UnhideRows(0, totalRows, -1);

        // Save the workbook as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("output.pdf", pdfOptions);
    }
}
