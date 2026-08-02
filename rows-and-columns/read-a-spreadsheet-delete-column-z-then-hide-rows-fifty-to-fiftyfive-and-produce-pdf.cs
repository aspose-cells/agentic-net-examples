// Title: C# – Delete Column Z, Hide Rows 50‑55 and Convert Excel to PDF with Aspose.Cells
// Description: Load an Excel workbook using Aspose.Cells for .NET, remove column Z, hide rows 50 through 55, and export the modified worksheet to a PDF file.
// Keywords: Aspose.Cells C# delete column | hide rows Aspose.Cells | Excel to PDF conversion .NET | remove column Z Aspose.Cells | Aspose.Cells PDF export
// Common Searches: how to delete a column and hide rows before PDF conversion using Aspose.Cells | Aspose.Cells C# hide rows 50 to 55 and export to PDF | remove column Z from Excel and save as PDF with Aspose.Cells
// Developer Intent: Delete column Z, hide rows 50‑55, then save the workbook as a PDF.
// Use Cases: Exclude confidential data column before generating a PDF report. | Suppress intermediate calculation rows when publishing a worksheet. | Prepare a clean Excel layout for PDF distribution by removing unwanted columns and rows.
// AI Prompts: Generate C# code with Aspose.Cells that deletes column Z, hides rows 50‑55, and saves the workbook as a PDF. | Explain the index calculations required to delete column Z and hide rows 50‑55 in Aspose.Cells before PDF export. | Show how to add error handling for missing worksheets when performing column deletion and row hiding with Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel workbook using Aspose.Cells for .NET, remove column Z, hide rows 50 through 55, and export the modified worksheet to a PDF file.
class Program
{
    static void Main()
    {
        // Paths for the source Excel file and the resulting PDF
        string inputFile = "input.xlsx";
        string outputFile = "output.pdf";

        // Load the workbook from the existing Excel file
        Workbook workbook = new Workbook(inputFile);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Delete column Z (zero‑based index 25)
        cells.DeleteColumn(25);

        // Hide rows 50 to 55 (zero‑based indices 49‑54, total 6 rows)
        cells.HideRows(49, 6);

        // Save the modified workbook as a PDF document
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}
