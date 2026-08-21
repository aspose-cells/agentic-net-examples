// Title: Aspose.Cells .NET: Delete rows 30‑42, insert 5 rows, and export to PDF
// Description: Loads an Excel workbook, removes rows 30‑42 from the first worksheet, inserts five blank rows at the same index with default height, and saves the modified file as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells delete rows | Aspose.Cells insert rows | Aspose.Cells export PDF | C# Excel row manipulation | remove specific rows Excel | default row height Aspose | save workbook as PDF
// Common Searches: how to delete a range of rows with Aspose.Cells .NET | insert rows with default height using Aspose.Cells | export worksheet to PDF after row operations | Aspose.Cells C# delete rows 30 to 42 | convert modified Excel to PDF with Aspose
// Developer Intent: Remove rows 30‑42, add five new rows at the same position, and generate a PDF of the workbook.
// Use Cases: Refresh a template by clearing old data rows, inserting placeholder rows, and producing a printable PDF report. | Re‑format a financial statement: cut out a section, add spacing rows, then distribute the final PDF to stakeholders. | Automate invoice preparation where specific rows are omitted, blank rows are added for notes, and the sheet is saved as a PDF.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete rows 30‑42, insert five rows at index 30, and save the workbook as a PDF. | Explain Aspose.Cells' zero‑based indexing for DeleteRows and InsertRows and how it maps to Excel row numbers. | Add comprehensive error handling for loading the source XLSX, performing row deletions/inserts, and writing the PDF output.

using System;
using Aspose.Cells;

namespace AsposeCellsRowManipulation
{
    // Loads an Excel workbook, removes rows 30‑42 from the first worksheet, inserts five blank rows at the same index with default height, and saves the modified file as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Delete rows 30 to 42 (inclusive). Row indices are zero‑based.
            // Total rows to delete = (42 - 30 + 1) = 13
            cells.DeleteRows(30, 13);

            // Insert five new rows at the same position (row index 30).
            // These rows will have the default height.
            cells.InsertRows(30, 5);

            // Export the modified workbook to PDF
            string outputPath = "output.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
