// Title: Hide Rows 5‑10 in an Excel Workbook and Export to PDF with Aspose.Cells for .NET
// Description: Load an existing Excel file using Aspose.Cells, hide rows 5‑10 on the first worksheet, and save the modified workbook directly as a PDF document.
// Keywords: Aspose.Cells hide rows C# | Aspose.Cells export PDF | C# hide Excel rows | Excel to PDF Aspose.Cells | hide rows before PDF conversion
// Common Searches: Aspose.Cells hide rows 5‑10 and save as PDF | C# hide specific rows in Excel before PDF export | How to hide rows in an Excel sheet with Aspose.Cells .NET | Export Excel to PDF after hiding rows using Aspose
// Developer Intent: Programmatically hide rows 5‑10 in the first worksheet and generate a PDF file.
// Use Cases: Create a printable PDF report that omits confidential or intermediate calculation rows. | Prepare a clean PDF version of a spreadsheet template for end‑user distribution. | Automate PDF generation from multiple worksheets while excluding hidden rows that contain draft data.
// AI Prompts: Generate C# code with Aspose.Cells that hides rows 5‑10 on every worksheet and exports the workbook to a single PDF with custom margins. | Explain how to conditionally hide rows based on cell values before converting an Excel file to PDF using Aspose.Cells for .NET. | Show how to hide rows, set PDF export options (e.g., fit to page, image quality), and save the result with Aspose.Cells in C#.

using System;
using Aspose.Cells;

// Load an existing Excel file using Aspose.Cells, hide rows 5‑10 on the first worksheet, and save the modified workbook directly as a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Hide rows 5 through 10 (zero‑based index 4, total 6 rows)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells.HideRows(4, 6);

        // Export the modified workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
