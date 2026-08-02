// Title: Hide rows 10‑20, unhide rows 15‑18 (auto‑fit) and save as PDF using Aspose.Cells for .NET
// Description: C# sample that loads an Excel workbook with Aspose.Cells, hides rows 10‑20, unhides rows 15‑18 applying automatic height, and converts the result to a PDF file.
// Keywords: Aspose.Cells | C# hide rows | Cells.HideRows | Cells.UnhideRows | auto‑fit row height | Excel to PDF conversion | export workbook as PDF | hide row range | unhide specific rows | .NET PDF generation
// Common Searches: Aspose.Cells hide rows 10 to 20 C# | unhide rows 15-18 auto fit height Aspose.Cells | convert Excel to PDF after hiding rows .NET | Cells.HideRows and Cells.UnhideRows example | how to hide and unhide rows in Aspose.Cells
// Developer Intent: Programmatically hide a set of rows, reveal a subset with automatic height adjustment, and generate a PDF from the modified workbook using C#.
// Use Cases: Create a printable report that omits confidential sections while displaying a specific range with proper spacing. | Prepare a PDF invoice where header rows are hidden during processing but later revealed for final output. | Automate document layout adjustments by toggling row visibility before exporting to PDF.
// AI Prompts: Write C# code with Aspose.Cells to hide rows 5‑12, then unhide rows 7‑9 using a fixed height of 20 points, and export to PDF. | Show how to use Cells.HideRows and Cells.UnhideRows with the auto‑fit height parameter (-1) in a .NET example. | Generate a step‑by‑step guide for hiding a row range, unhiding a sub‑range with auto‑fit, and converting the worksheet to PDF using Aspose.Cells.

using System;
using Aspose.Cells;

// C# sample that loads an Excel workbook with Aspose.Cells, hides rows 10‑20, unhides rows 15‑18 applying automatic height, and converts the result to a PDF file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide rows 10 through 20 (zero‑based index 9, total 11 rows)
        cells.HideRows(9, 11);

        // Unhide rows 15 through 18 (zero‑based index 14, total 4 rows) with auto‑fit height (-1)
        cells.UnhideRows(14, 4, -1);

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
