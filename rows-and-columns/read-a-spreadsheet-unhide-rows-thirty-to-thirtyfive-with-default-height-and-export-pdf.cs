// Title: C# – Unhide Rows 30‑35 and Export Excel to PDF with Aspose.Cells
// Description: Loads an existing workbook (input.xlsx) with Aspose.Cells, unhides rows 30‑35 on the first worksheet using Cells.UnhideRows (default height ‑1), and saves the result as a PDF (output.pdf).
// Keywords: Aspose.Cells C# unhide rows | Cells.UnhideRows | Excel to PDF conversion .NET | default row height -1 | batch Excel PDF Aspose | unhide hidden rows Aspose.Cells
// Common Searches: Aspose.Cells how to unhide specific rows | C# unhide rows 30-35 Excel | Convert Excel to PDF after unhiding rows Aspose | Set default row height when unhiding rows Aspose.Cells | Export hidden rows to PDF using Aspose.Cells
// Developer Intent: Reveal rows 30‑35 in an Excel sheet and generate a PDF version.
// Use Cases: Create printable reports where hidden rows must appear. | Automate bulk workbook processing to ensure rows 30‑35 are visible before PDF conversion. | Maintain consistent row spacing by resetting hidden rows to default height during export. | Prepare financial statements that hide intermediate rows in the source but require full visibility in the final PDF.
// AI Prompts: Generate C# code that uses Aspose.Cells to unhide rows 30‑35 with default height and save the workbook as a PDF. | Explain the meaning of the parameters in Cells.UnhideRows, especially the -1 height value. | Show how to add error handling for file loading, row unhiding, and PDF saving with Aspose.Cells. | Provide a sample that processes multiple Excel files, unhides rows 30‑35, and converts each to PDF. | Demonstrate how to verify that rows were successfully unhidden before exporting.

using System;
using Aspose.Cells;

// Loads an existing workbook (input.xlsx) with Aspose.Cells, unhides rows 30‑35 on the first worksheet using Cells.UnhideRows (default height ‑1), and saves the result as a PDF (output.pdf).
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Unhide rows 30 to 35 (zero‑based index) with default height (-1)
        // totalRows = 6 because rows 30,31,32,33,34,35 are to be unhidden
        workbook.Worksheets[0].Cells.UnhideRows(30, 6, -1);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
