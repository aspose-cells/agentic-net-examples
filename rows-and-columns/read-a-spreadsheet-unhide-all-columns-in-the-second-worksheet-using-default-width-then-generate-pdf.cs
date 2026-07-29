// Title: C# – Unhide All Columns in the Second Worksheet (Default Width) and Save as PDF with Aspose.Cells
// Description: Load an Excel workbook, target the second worksheet, reveal every hidden column using the default width (-1), and convert the sheet to a PDF file with optional PdfSaveOptions.
// Keywords: Aspose.Cells C# unhide columns | second worksheet PDF export | default column width Aspose | unhide columns before PDF conversion | Aspose.Cells PDFSaveOptions
// Common Searches: how to unhide columns in a specific sheet using Aspose.Cells | Aspose.Cells hide columns then export to PDF | C# unhide all columns default width Aspose | convert Excel sheet to PDF after unhiding columns
// Developer Intent: Reveal every column in the second worksheet with the default width and generate a PDF from the workbook.
// Use Cases: Creating PDF reports where hidden columns must be visible for compliance. | Batch processing of workbooks to ensure consistent column visibility before distribution. | Automating document pipelines that require default‑width column layout in the final PDF.
// AI Prompts: Generate C# code that opens an Excel file, unhides all columns in worksheet index 1 using the default width, and saves the result as a PDF. | Write a reusable Aspose.Cells method that accepts a file path and sheet index, unhides columns with width -1, and returns a PDF byte array. | Provide C# error‑handling examples for missing worksheets or invalid file paths when unhiding columns and exporting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an Excel workbook, target the second worksheet, reveal every hidden column using the default width (-1), and convert the sheet to a PDF file with optional PdfSaveOptions.
class UnhideColumnsAndConvertToPdf
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1, zero‑based)
        Worksheet sheet = workbook.Worksheets[1];

        // Determine the total number of columns in the worksheet
        int totalColumns = sheet.Cells.Columns.Count;

        // Unhide all columns using the default column width (-1)
        sheet.Cells.UnhideColumns(0, totalColumns, -1);

        // Prepare PDF save options (optional, can be omitted for default settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
