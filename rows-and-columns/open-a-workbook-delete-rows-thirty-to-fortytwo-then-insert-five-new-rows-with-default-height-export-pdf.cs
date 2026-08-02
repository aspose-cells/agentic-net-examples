// Title: C# – Delete rows 30‑42, insert 5 rows, and export to PDF with Aspose.Cells
// Description: A concise C# example that loads an Excel workbook, removes rows 30‑42 from the first worksheet, inserts five new rows with the default height, and saves the result directly as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | delete rows | insert rows | Excel to PDF | row manipulation | Workbook.Save | Cells.DeleteRows | Cells.InsertRows | PDF export
// Common Searches: Aspose.Cells delete rows 30 to 42 C# | Insert rows before PDF export Aspose.Cells | How to remove a range of rows and save as PDF with Aspose.Cells | C# code to delete and insert rows then export Excel to PDF | Aspose.Cells row range manipulation example
// Developer Intent: Remove a specific block of rows, add a set of new rows at the same location, and generate a PDF from the updated worksheet.
// Use Cases: Refresh a report template by deleting outdated rows, inserting placeholder rows for new data, and delivering the final layout as a PDF to clients. | Automate archival of Excel sheets where certain rows are stripped out, fresh rows are added for summary information, and the document is saved in PDF format for compliance.
// AI Prompts: Generate C# code with Aspose.Cells that deletes rows 30‑42, inserts five rows, and saves the workbook as a PDF. | Explain how to control row height after inserting rows with Aspose.Cells before converting the sheet to PDF. | Provide a step‑by‑step tutorial for removing a row range, adding new rows, and exporting the worksheet to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRowManipulation
{
    // A concise C# example that loads an Excel workbook, removes rows 30‑42 from the first worksheet, inserts five new rows with the default height, and saves the result directly as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook (lifecycle rule: use load)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Delete rows 30 to 42 (zero‑based index: start at 29, delete 13 rows)
            int startRowToDelete = 29;   // Row 30
            int rowsToDelete = 13;       // Rows 30‑42 inclusive
            cells.DeleteRows(startRowToDelete, rowsToDelete);

            // Insert five new rows at the same position (default height is applied automatically)
            int rowsToInsert = 5;
            cells.InsertRows(startRowToDelete, rowsToInsert);

            // Save the modified workbook as PDF (lifecycle rule: use save)
            string outputPath = "output.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
