// Title: C# – Hide rows 21‑25 in an Excel sheet and export to PDF using Aspose.Cells
// Description: Load an existing workbook with Aspose.Cells for .NET, programmatically conceal rows 21‑25 via the zero‑based HideRows method, and save the result directly as a PDF file.
// Keywords: Aspose.Cells | C# | .NET | HideRows | hide rows | Excel to PDF | export PDF | row visibility | programmatic Excel | PDF conversion | input.xlsx | output.pdf
// Common Searches: Aspose.Cells hide rows 21 to 25 C# | Export Excel worksheet to PDF after hiding rows | How to use HideRows method Aspose.Cells .NET | Convert Excel to PDF with hidden rows using C# | Hide specific rows before PDF conversion Aspose
// Developer Intent: Programmatically conceal rows 21‑25 in an Excel workbook and generate a PDF of the worksheet.
// Use Cases: Create a clean printable PDF that omits confidential or intermediate data rows. | Produce a summary report where detailed rows are hidden to highlight key metrics. | Automate document generation for client‑facing PDFs while preserving the original workbook layout.
// AI Prompts: Generate C# code that uses Aspose.Cells to hide rows 21‑25 and save the worksheet as a PDF. | Explain the zero‑based indexing of the HideRows method and how to adjust it for different row ranges. | Add robust error handling for missing files and permission issues when converting an Excel file to PDF after hiding rows.

using System;
using Aspose.Cells;

// Load an existing workbook with Aspose.Cells for .NET, programmatically conceal rows 21‑25 via the zero‑based HideRows method, and save the result directly as a PDF file.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 21 through 25 (zero‑based index 20, total 5 rows)
        worksheet.Cells.HideRows(20, 5);

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
