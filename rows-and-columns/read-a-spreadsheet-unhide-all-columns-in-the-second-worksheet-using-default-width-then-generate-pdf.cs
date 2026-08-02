// Title: C# – Unhide All Columns in Second Worksheet and Export to PDF with Aspose.Cells
// Description: Loads an Excel file, accesses the second worksheet, calculates the last used column, unhides every column using the default width (Width = -1), and saves the workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | unhide columns | second worksheet | PDF export | default column width | Cells.UnhideColumns | Excel to PDF | .NET | PdfSaveOptions | workbook.Save
// Common Searches: Aspose.Cells unhide columns second sheet C# | export Excel worksheet to PDF with hidden columns visible | C# code to unhide all columns before PDF conversion | Cells.UnhideColumns default width example | convert Excel to PDF using Aspose.Cells .NET
// Developer Intent: Reveal every column in the second worksheet and generate a PDF of the workbook.
// Use Cases: Create PDF reports where hidden columns must be displayed for full data visibility. | Automate batch processing of Excel files to ensure all columns appear in the final PDF output. | Standardize column widths to the default before converting worksheets to PDF for consistent layout.
// AI Prompts: Generate C# code that uses Aspose.Cells to unhide all columns in the second worksheet of an Excel file and then save the workbook as a PDF. | Show how to determine the last used column in a worksheet, call Cells.UnhideColumns with Width = -1, and export the result to PDF using PdfSaveOptions. | Explain the steps to load an Excel workbook, unhide hidden columns on a specific sheet, and convert the workbook to PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel file, accesses the second worksheet, calculates the last used column, unhides every column using the default width (Width = -1), and saves the workbook as a PDF using Aspose.Cells for .NET.
class UnhideColumnsAndExportPdf
{
    static void Main()
    {
        // Input Excel file path
        string inputFile = "input.xlsx";

        // Output PDF file path
        string outputFile = "output.pdf";

        // Load the workbook from the existing Excel file
        Workbook workbook = new Workbook(inputFile);

        // Access the second worksheet (index 1, zero‑based)
        Worksheet worksheet = workbook.Worksheets[1];
        Cells cells = worksheet.Cells;

        // Determine the total number of columns that contain data.
        // MaxColumn returns the zero‑based index of the last column with data.
        int lastColumnIndex = cells.MaxColumn;
        int totalColumns = lastColumnIndex + 1; // convert to count

        // Unhide all columns in the worksheet using the default column width.
        // Width = -1 tells Aspose.Cells to apply the standard width.
        cells.UnhideColumns(0, totalColumns, -1);

        // Save the modified workbook as a PDF document.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save(outputFile, pdfOptions);
    }
}
