// Title: C# – Delete rows 30‑42, insert 5 rows, and save as PDF with Aspose.Cells
// Description: Loads an Excel workbook, removes rows 30‑42 (13 rows), inserts five blank rows at the same location with default height, and converts the result to a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# delete rows | Aspose.Cells insert rows | Aspose.Cells PDF export | Excel row manipulation .NET | remove rows Excel Aspose | add rows before PDF conversion
// Common Searches: Aspose.Cells delete rows 30 to 42 C# | Insert rows after deleting a range with Aspose.Cells | Export modified Excel to PDF using Aspose.Cells | C# code to replace specific rows in an Excel file | How to remove and add rows before PDF conversion in .NET
// Developer Intent: Remove rows 30‑42, add five new rows at that position, and generate a PDF from the updated workbook.
// Use Cases: Prepare a printable report by deleting a section of data and inserting placeholder rows to keep layout consistency. | Automate preprocessing of uploaded Excel files where a specific row block must be replaced before PDF generation. | Create a clean PDF version of a spreadsheet after programmatically adjusting its row structure.
// AI Prompts: Generate C# code with Aspose.Cells that deletes rows 30‑42, inserts five rows, and saves the file as PDF. | Explain how DeleteRows and InsertRows use zero‑based indexing in Aspose.Cells. | Add robust error handling for missing input files and PDF save failures in the row‑manipulation example.

using System;
using Aspose.Cells;

namespace AsposeCellsRowManipulation
{
    // Loads an Excel workbook, removes rows 30‑42 (13 rows), inserts five blank rows at the same location with default height, and converts the result to a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Delete rows 30 to 42 (zero‑based index 29, total 13 rows)
            cells.DeleteRows(29, 13);

            // Insert five new rows at the same position (default height)
            cells.InsertRows(29, 5);

            // Export the modified workbook to PDF (save rule)
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
