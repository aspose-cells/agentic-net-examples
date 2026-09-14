// Title: Unhide every column in the second worksheet of an Excel workbook and save it as PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX file, makes all columns visible in the worksheet at index 1 with default width, and then saves the workbook as a PDF using Aspose.Cells. | Show how to programmatically call the UnhideColumns method for the entire column range of the second sheet before exporting to PDF with Aspose.Cells in C#. | Provide a step‑by‑step example that reads input.xlsx, reveals hidden columns on the second worksheet, and generates output.pdf using Aspose.Cells Rendering.
// Common Searches: Aspose.Cells C# unhide all columns on a specific sheet before PDF conversion | how to export second worksheet of Excel to PDF after making hidden columns visible in .NET | C# code to unhide columns with default width using Aspose.Cells UnhideColumns method | convert XLSX to PDF with all columns shown in sheet index 1 using Aspose.Cells | unhide columns programmatically in Aspose.Cells and save workbook as PDF
// Tags: Aspose.Cells unhide columns C# | export second sheet to PDF Aspose.Cells | UnhideColumns default width usage | Excel workbook PDF conversion with visible columns | programmatic column visibility Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads input.xlsx, unhides all columns in the second worksheet using the default width, and saves the result as output.pdf.
class UnhideColumnsAndConvertToPdf
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1)
        Worksheet worksheet = workbook.Worksheets[1];

        // Determine the number of columns to unhide.
        // MaxColumn returns the last column index that contains data, so add 1 for count.
        int totalColumns = worksheet.Cells.MaxColumn + 1;

        // Unhide all columns starting from index 0 with default width (-1).
        worksheet.Cells.UnhideColumns(0, totalColumns, -1);

        // Save the workbook as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("output.pdf", pdfOptions);
    }
}
