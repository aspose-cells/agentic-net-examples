// Title: Hide Rows 5‑10 in an Excel Workbook and Export to PDF using Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to load an existing Excel file with Aspose.Cells, hide rows 5 through 10 on the first worksheet using the zero‑based HideRows method, and then save the modified workbook as a PDF with default PdfSaveOptions.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | HideRows | hide rows 5 to 10 | Excel to PDF | export workbook as PDF | row visibility | PdfSaveOptions | programmatic PDF conversion
// Common Searches: Aspose.Cells hide rows 5‑10 C# | Excel hide rows before PDF conversion Aspose | How to export hidden rows Excel to PDF using Aspose.Cells | C# hide specific rows in worksheet and save as PDF | Aspose.Cells HideRows method example
// Developer Intent: Hide rows 5‑10 in an Excel worksheet and save the workbook as a PDF.
// Use Cases: Create printable PDFs that exclude confidential or intermediate calculation rows. | Automate report generation where only selected rows should appear in the final PDF. | Prepare archival PDFs of spreadsheets after programmatically removing unwanted rows.
// AI Prompts: Generate C# code with Aspose.Cells to hide rows 5‑10 on the first worksheet and export the workbook to PDF. | Explain how the HideRows method uses zero‑based indexing and how to combine it with PdfSaveOptions for PDF output. | Show a step‑by‑step tutorial for loading an Excel file, hiding specific rows, and saving as PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// This example demonstrates how to load an existing Excel file with Aspose.Cells, hide rows 5 through 10 on the first worksheet using the zero‑based HideRows method, and then save the modified workbook as a PDF with default PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Hide rows 5 through 10 (zero‑based index starts at 4, total rows = 6)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells.HideRows(4, 6);

        // Prepare PDF save options (default options are sufficient)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the modified workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
