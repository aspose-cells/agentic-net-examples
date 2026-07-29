// Title: C# – Remove column Q with Aspose.Cells and export the sheet to PDF
// Description: Loads an Excel workbook, deletes column Q (zero‑based index 16) from the first worksheet, and saves the result directly as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | remove column | column Q | zero‑based index | Excel to PDF conversion | worksheet column deletion | save as PDF | Aspose.Cells API | C# Excel manipulation | PDF export .NET
// Common Searches: Aspose.Cells remove column Q C# | How to delete a specific column before PDF export | Convert Excel to PDF after column removal | C# code to delete column by index Aspose | Export modified worksheet to PDF Aspose.Cells
// Developer Intent: Strip column Q from an Excel file and generate a PDF of the cleaned worksheet.
// Use Cases: Redact sensitive identifiers before sharing a PDF report. | Eliminate calculation columns when creating a presentation‑ready document. | Automate preprocessing of uploaded spreadsheets prior to archiving them as PDFs.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete column Q (index 16) from the first worksheet and save the workbook as a PDF. | Explain how to map an Excel column letter to its zero‑based index in Aspose.Cells and then perform a PDF export. | Provide a sample that removes multiple columns and creates separate PDF files for each worksheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads an Excel workbook, deletes column Q (zero‑based index 16) from the first worksheet, and saves the result directly as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the source workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete column Q.
        // Column indexes are zero‑based, so column Q corresponds to index 16.
        sheet.Cells.DeleteColumn(16);

        // Save the modified worksheet as a PDF document.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
