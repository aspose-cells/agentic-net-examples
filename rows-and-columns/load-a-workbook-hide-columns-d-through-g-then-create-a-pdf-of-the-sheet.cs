// Title: C# – Hide Columns D‑G in Excel and Export to PDF with Aspose.Cells
// Description: Load an Excel workbook, hide columns D through G using Cells.HideColumns, and save the worksheet as a PDF with SaveFormat.Pdf in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | HideColumns | C# | Excel to PDF | hide columns D G | column visibility | worksheet PDF conversion | Aspose.Cells .NET
// Common Searches: Aspose.Cells hide columns D to G | C# hide multiple columns before PDF export | How to hide Excel columns with Aspose.Cells | Export hidden columns to PDF using Aspose.Cells | Cells.HideColumns example C#
// Developer Intent: Hide specific columns in a worksheet and generate a PDF of the modified sheet.
// Use Cases: Remove confidential data columns before sharing a PDF report. | Create printable PDFs that display only the required data columns. | Produce different PDF versions of the same workbook for distinct audiences by toggling column visibility.
// AI Prompts: Generate C# code that uses Aspose.Cells to hide columns D through G in an Excel file and save the result as a PDF. | Explain the zero‑based indexing used by Cells.HideColumns in Aspose.Cells and how it affects PDF export. | Provide a step‑by‑step tutorial for programmatically hiding a range of columns and converting the worksheet to PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an Excel workbook, hide columns D through G using Cells.HideColumns, and save the worksheet as a PDF with SaveFormat.Pdf in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns D (index 3) through G (index 6)
        // Total columns to hide = 4 (D, E, F, G)
        int startColumn = 3; // Column D (zero‑based)
        int totalColumns = 4; // D, E, F, G
        cells.HideColumns(startColumn, totalColumns);

        // Save the modified workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
