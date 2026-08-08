// Title: Hide Rows 21‑25 in Excel and Export to PDF with Aspose.Cells for .NET
// Description: Load an Excel workbook, hide rows 21‑25 using Cells.HideRows (zero‑based start index 20, count 5), and save the worksheet as a PDF. The hidden rows are omitted from the generated PDF view.
// Keywords: Aspose.Cells | C# hide rows | HideRows method | Excel to PDF conversion | exclude rows from PDF | Aspose.Cells PDF export
// Common Searches: Aspose.Cells hide specific rows C# | Export Excel to PDF without certain rows | Hide rows 21 to 25 using Aspose.Cells | C# hide rows and save as PDF | How to exclude rows from PDF in Aspose.Cells
// Developer Intent: Hide selected rows in an Excel worksheet and generate a PDF.
// Use Cases: Create a client‑ready PDF that omits confidential rows. | Produce clean printable reports by removing internal data before export. | Automate batch conversion of workbooks where designated rows must be hidden in the PDF output.
// AI Prompts: Generate C# code with Aspose.Cells to hide rows 30‑35 and save the sheet as a PDF. | Explain how zero‑based indexing works for the HideRows method and how to calculate the start index. | Add error handling to the example for missing input files and PDF write failures.

using System;
using Aspose.Cells;

namespace AsposeCellsHideRowsAndSavePdf
{
    // Load an Excel workbook, hide rows 21‑25 using Cells.HideRows (zero‑based start index 20, count 5), and save the worksheet as a PDF. The hidden rows are omitted from the generated PDF view.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can change the index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide rows 21 to 25.
            // Row indices are zero‑based, so row 21 is index 20.
            // HideRows(startRowIndex, totalRows) hides a consecutive block of rows.
            worksheet.Cells.HideRows(20, 5);

            // Save the workbook as PDF. The hidden rows will be reflected in the PDF view.
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
