// Title: Remove Row 8 from an Excel workbook and export to PDF with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file using Aspose.Cells, deletes the eighth row (zero‑based index 7) from the first worksheet, and saves the modified workbook directly as a PDF document.
// Keywords: Aspose.Cells DeleteRow C# | remove Excel row .NET | Excel to PDF conversion Aspose | Workbook.Save PDF Aspose.Cells | delete specific row worksheet | C# Excel manipulation PDF export | Aspose.Cells row removal example
// Common Searches: Aspose.Cells delete row 8 C# | convert Excel to PDF after removing a row | C# code to delete a row in Excel and save as PDF | Aspose.Cells remove specific row before PDF export | how to delete a worksheet row with Aspose.Cells
// Developer Intent: Delete the eighth row of an Excel sheet and generate a PDF version of the updated workbook.
// Use Cases: Strip a header or placeholder row before creating a client‑ready PDF report. | Omit confidential or temporary data rows when publishing spreadsheets as PDFs. | Automate cleanup of generated Excel files in a batch process prior to PDF conversion.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete a given row number from an Excel worksheet and then saves the workbook as a PDF. | Explain how to validate the row index before calling DeleteRow to avoid out‑of‑range errors in Aspose.Cells. | Show an example that removes multiple consecutive rows and exports the result to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads an Excel file using Aspose.Cells, deletes the eighth row (zero‑based index 7) from the first worksheet, and saves the modified workbook directly as a PDF document.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Delete row 8 (zero‑based index 7) in the first worksheet
        workbook.Worksheets[0].Cells.DeleteRow(7);

        // Save the updated workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
