// Title: C# – Unhide All Rows (default height) in First Sheet and Export to PDF with Aspose.Cells
// Description: Load an Excel workbook, reveal every hidden row in the first worksheet using UnhideRows with a height of -1 (auto‑fit), and then save the file as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | UnhideRows | default row height | export to PDF | Excel to PDF conversion | reveal hidden rows | first worksheet | SaveFormat.Pdf | .NET example
// Common Searches: Aspose.Cells unhide all rows C# | How to export Excel to PDF after unhiding rows Aspose.Cells | UnhideRows height -1 Aspose.Cells | Convert hidden rows Excel to PDF .NET | Reveal hidden rows in first sheet Aspose.Cells
// Developer Intent: Make every row in the first worksheet visible with its default (auto‑fit) height and generate a PDF version of the workbook.
// Use Cases: Create PDF reports from templates that may contain hidden rows, ensuring all data appears with standard row heights. | Automate batch processing of workbooks where hidden rows must be shown before converting each file to PDF. | Build a web service that accepts an uploaded Excel file, unhides rows in the first sheet, and returns a PDF for downstream consumption.
// AI Prompts: Generate C# code that uses Aspose.Cells to unhide all rows in the first worksheet with default height and save the workbook as a PDF. | Explain the effect of passing -1 as the height argument to the UnhideRows method in Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for loading an Excel file, revealing hidden rows, and exporting the result to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel workbook, reveal every hidden row in the first worksheet using UnhideRows with a height of -1 (auto‑fit), and then save the file as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Unhide all rows in the first sheet.
        // The third parameter (-1) tells Aspose.Cells to use the default (auto‑fit) height.
        // Using a large row count ensures all possible rows are processed.
        sheet.Cells.UnhideRows(0, sheet.Cells.Rows.Count, -1);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
