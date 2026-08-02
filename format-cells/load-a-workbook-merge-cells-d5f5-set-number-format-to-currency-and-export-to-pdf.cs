// Title: C# – Merge D5:F5, set currency format, and export to PDF with Aspose.Cells
// Description: Load an Excel workbook, merge the range D5:F5 on the first worksheet, apply the built‑in currency number format (ID 164), and save the file as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# merge cells | currency number format Aspose.Cells | export Excel to PDF .NET | built‑in format 164 | merge D5 F5 | Excel to PDF conversion | Aspose.Cells sample code | C# Excel formatting PDF
// Common Searches: Aspose.Cells merge D5:F5 C# | set currency format for merged cells Aspose | export formatted Excel workbook to PDF .NET | how to apply built‑in number format Aspose.Cells | C# code to merge cells and save as PDF
// Developer Intent: Merge cells D5:F5, apply a currency number format, and generate a PDF from the workbook.
// Use Cases: Create a printable financial summary where the total amount cell spans D5:F5 and shows currency values. | Generate invoices that merge the amount column, format it as currency, and deliver the final document as a PDF. | Produce a sales dashboard PDF with merged header cells styled in a currency format for easy distribution.
// AI Prompts: Write C# code with Aspose.Cells to merge D5:F5, set the built‑in currency format (ID 164), and save the workbook as a PDF. | Explain how to apply a built‑in currency number format to a merged cell range before exporting to PDF using Aspose.Cells for .NET. | Provide step‑by‑step instructions for merging a cell range, formatting it as currency, and converting the Excel file to PDF with Aspose.Cells.

using Aspose.Cells;
using System;

// Load an Excel workbook, merge the range D5:F5 on the first worksheet, apply the built‑in currency number format (ID 164), and save the file as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells D5:F5.
        // Row and column indices are zero‑based, so D5 is row 4, column 3.
        cells.Merge(4, 3, 1, 3);

        // Set the number format of the merged cell to a built‑in currency format.
        Style style = cells[4, 3].GetStyle();
        style.Number = 164; // Built‑in currency format (e.g., $#,##0.00)
        cells[4, 3].SetStyle(style);

        // Export the workbook to PDF. The format is inferred from the file extension.
        workbook.Save("output.pdf");
    }
}
