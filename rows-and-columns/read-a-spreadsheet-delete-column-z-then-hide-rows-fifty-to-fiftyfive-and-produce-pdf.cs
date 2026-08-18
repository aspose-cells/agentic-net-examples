// Title: C# – Delete Column Z, Hide Rows 50‑55 and Export Excel to PDF with Aspose.Cells
// Description: Loads an XLSX workbook, removes column Z (index 25), hides rows 50‑55 (zero‑based start 49, count 6), and saves the result directly as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | delete column Z | hide rows 50-55 | Excel to PDF conversion | worksheet manipulation | zero based indexing | .NET Excel export | remove specific column | hide multiple rows
// Common Searches: Aspose.Cells delete column Z C# | How to hide rows 50 to 55 with Aspose.Cells | Export modified Excel workbook to PDF using Aspose.Cells .NET | C# code to remove a column and hide rows before PDF conversion | Aspose.Cells hide rows range and save as PDF
// Developer Intent: Delete column Z, hide rows 50‑55, then save the workbook as a PDF.
// Use Cases: Prepare a client‑facing financial report by stripping confidential column Z and concealing internal calculation rows before generating a PDF. | Create a clean printable invoice where internal notes in column Z are removed and detailed line items (rows 50‑55) are hidden, then export to PDF for distribution. | Produce a sanitized spreadsheet for external partners by deleting a specific column, hiding auxiliary rows, and converting the worksheet to PDF in a single automated step.
// AI Prompts: Generate C# code using Aspose.Cells to delete column Z, hide rows 50‑55, and save the workbook as a PDF. | Show an Aspose.Cells example that removes a specific column and hides a range of rows before exporting to PDF, including zero‑based index details. | Explain step‑by‑step how to modify an Excel worksheet with Aspose.Cells: delete column Z, hide rows 50‑55, then convert the sheet to PDF.

using System;
using Aspose.Cells;

// Loads an XLSX workbook, removes column Z (index 25), hides rows 50‑55 (zero‑based start 49, count 6), and saves the result directly as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Delete column Z (zero‑based index 25)
        cells.DeleteColumn(25);

        // Hide rows 50‑55 (zero‑based start index 49, total 6 rows)
        cells.HideRows(49, 6);

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
