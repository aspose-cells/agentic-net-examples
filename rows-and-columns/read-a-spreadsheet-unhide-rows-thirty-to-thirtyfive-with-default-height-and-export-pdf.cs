// Title: C# – Unhide rows 30‑35 and export Excel to PDF with Aspose.Cells
// Description: Loads an Excel workbook, unhides rows 30‑35 using the default (auto‑fit) height, and saves the file as a PDF via Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | UnhideRows | default row height | export to PDF | Excel to PDF | row visibility | worksheet manipulation | batch spreadsheet processing
// Common Searches: Aspose.Cells unhide rows 30 to 35 C# | How to export hidden rows to PDF using Aspose.Cells | Set auto‑fit row height when unhiding rows Aspose.Cells | C# code to unhide multiple rows and save as PDF | Aspose.Cells UnhideRows method example
// Developer Intent: Reveal rows 30‑35 in an Excel sheet and generate a PDF of the workbook.
// Use Cases: Create printable PDF reports where previously hidden rows must be visible. | Automate batch conversion of spreadsheets, ensuring all rows appear in the final PDF. | Prepare a PDF version of a template after programmatically adjusting row visibility for specific sections.
// AI Prompts: Generate C# code that unhides rows 30‑35 with default height using Aspose.Cells and saves the workbook as a PDF. | Explain the parameters of the UnhideRows method in Aspose.Cells, especially the use of -1 for auto‑fit row height. | Show error‑handling patterns for loading an Excel file, unhiding rows, and exporting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook, unhides rows 30‑35 using the default (auto‑fit) height, and saves the file as a PDF via Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide rows 30 to 35 (zero‑based index) with default (auto‑fit) height
        // totalRows = 6 because rows 30,31,32,33,34,35 are to be unhidden
        worksheet.Cells.UnhideRows(30, 6, -1);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
