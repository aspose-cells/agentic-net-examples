// Title: C# – Hide Columns 5‑8, Unhide 6‑7 with Width, Export to PDF using Aspose.Cells
// Description: Loads a workbook with Aspose.Cells for .NET, hides columns E‑H, unhides columns F‑G while setting a 15‑point width, and saves the worksheet as a PDF file.
// Keywords: Aspose.Cells | C# | .NET spreadsheet | hide columns | unhide columns | set column width | export to PDF | worksheet PDF conversion | column visibility | Aspose.Cells PDF export
// Common Searches: Aspose.Cells hide columns C# | unhide columns with width Aspose.Cells | export worksheet to PDF Aspose.Cells | set column width after unhiding Aspose.Cells | C# hide columns E to H Aspose.Cells
// Developer Intent: Hide columns 5‑8, unhide columns 6‑7 with a 15‑point width, then export the sheet as a PDF.
// Use Cases: Create a printable PDF where sensitive columns are hidden but key columns retain a custom width. | Generate a report that omits intermediate columns while preserving layout for selected columns. | Prepare a PDF export after adjusting column visibility and widths to meet specific formatting standards.
// AI Prompts: Write C# code with Aspose.Cells to hide columns E‑H, unhide columns F‑G with a width of 15 points, and save the workbook as a PDF. | Show an Aspose.Cells example that demonstrates hiding a range of columns, setting the width of selected columns, and exporting the worksheet to PDF.

using System;
using Aspose.Cells;

// Loads a workbook with Aspose.Cells for .NET, hides columns E‑H, unhides columns F‑G while setting a 15‑point width, and saves the worksheet as a PDF file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns 5 through 8 (zero‑based indices 4‑7)
        int startHideColumn = 4;      // Column E (5th column)
        int hideColumnCount = 4;      // Columns E, F, G, H
        cells.HideColumns(startHideColumn, hideColumnCount);

        // Unhide columns 6 and 7 (zero‑based indices 5‑6) and set their width
        int startUnhideColumn = 5;    // Column F (6th column)
        int unhideColumnCount = 2;    // Columns F and G
        double columnWidth = 15.0;    // Desired width for the unhidden columns
        cells.UnhideColumns(startUnhideColumn, unhideColumnCount, columnWidth);

        // Export the worksheet to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
