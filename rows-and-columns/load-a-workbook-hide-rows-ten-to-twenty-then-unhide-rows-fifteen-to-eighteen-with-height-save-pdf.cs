// Title: Hide rows 10‑20, unhide rows 15‑18 (auto‑fit), and export to PDF with Aspose.Cells for .NET
// Description: Load an Excel workbook, hide rows 10‑20, then unhide rows 15‑18 while automatically fitting their height, and save the modified file as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | HideRows | UnhideRows | auto‑fit row height | export to PDF | Excel row visibility | Cells.HideRows | Cells.UnhideRows | PDF conversion | Aspose.Cells .NET
// Common Searches: Aspose.Cells hide rows 10 to 20 | unhide rows 15 to 18 auto fit height Aspose | export Excel to PDF after hiding rows C# | how to hide and unhide rows with Aspose.Cells | row visibility PDF conversion Aspose.Cells
// Developer Intent: Hide rows 10‑20, unhide rows 15‑18 with auto‑fit, and generate a PDF.
// Use Cases: Create a printable report that collapses detailed sections (rows 10‑20) but shows a summary block (rows 15‑18) with proper height before PDF export. | Prepare a financial statement where intermediate line items are hidden, then reveal total rows with auto‑adjusted height to maintain layout in the final PDF. | Generate an invoice PDF that hides line‑item rows while displaying the totals area with automatically fitted row height.
// AI Prompts: Write C# code using Aspose.Cells to hide rows 5‑12, then unhide rows 8‑10 with auto‑fit height, and save the workbook as a PDF. | Show an example of Cells.HideRows and Cells.UnhideRows with the height parameter set to -1 to adjust row height before exporting to PDF.

using System;
using Aspose.Cells;

// Load an Excel workbook, hide rows 10‑20, then unhide rows 15‑18 while automatically fitting their height, and save the modified file as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide rows 10 through 20 (zero‑based index 9, total 11 rows)
        cells.HideRows(9, 11);

        // Unhide rows 15 through 18 (zero‑based index 14, total 4 rows) with auto‑fit height (-1)
        cells.UnhideRows(14, 4, -1);

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
