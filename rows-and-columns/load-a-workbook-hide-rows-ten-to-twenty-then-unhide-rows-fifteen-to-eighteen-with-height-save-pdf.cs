// Title: Hide rows 10‑20, unhide rows 15‑18 with auto‑fit height, and export to PDF using Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, hides rows 10‑20, reveals rows 15‑18 while automatically adjusting their height, and saves the result as a PDF document with Aspose.Cells for C#.
// Keywords: Aspose.Cells hide rows C# | Aspose.Cells UnhideRows method | auto‑fit row height Aspose.Cells | export worksheet to PDF C# | row visibility manipulation Aspose.Cells | HideRows API Aspose.Cells | C# Excel to PDF conversion
// Common Searches: Aspose.Cells hide a range of rows then unhide a subset | C# unhide rows with automatic height using Aspose.Cells | How to export a workbook to PDF after changing row visibility | Aspose.Cells HideRows and UnhideRows example | Set row height to auto when unhiding rows in Aspose.Cells
// Developer Intent: The developer needs to conceal rows 10‑20, make rows 15‑18 visible with auto‑adjusted height, and generate a PDF from the modified worksheet.
// Use Cases: Create a printable report that initially hides confidential sections, then reveals only the required rows with proper spacing before PDF export. | Generate invoices where summary rows are hidden during processing but line‑item rows are displayed with optimal height in the final PDF. | Prepare a client‑ready spreadsheet where business rules dictate which rows stay hidden and which are shown with auto‑fit height for clean PDF output.
// AI Prompts: Provide C# code that hides rows 10‑20, unhides rows 15‑18 with auto‑fit height, and saves the workbook as a PDF using Aspose.Cells. | Explain the effect of passing -1 as the height parameter to the UnhideRows method in Aspose.Cells. | Show an example of converting an Excel worksheet to PDF after modifying row visibility with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRowVisibilityDemo
{
    // Loads an existing workbook, hides rows 10‑20, reveals rows 15‑18 while automatically adjusting their height, and saves the result as a PDF document with Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Hide rows 10 to 20 (zero‑based index 9, total 11 rows)
            cells.HideRows(9, 11);

            // Unhide rows 15 to 18 (zero‑based index 14, total 4 rows) and auto‑fit height
            cells.UnhideRows(14, 4, -1);

            // Save the modified workbook as PDF
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
