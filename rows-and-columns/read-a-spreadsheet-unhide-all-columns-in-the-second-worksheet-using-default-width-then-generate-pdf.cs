// Title: Unhide All Columns in the Second Worksheet and Export to PDF with Aspose.Cells for .NET
// Description: Load an Excel workbook, access the second worksheet, unhide every column using the default width (‑1), and save the result as a PDF file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unhide columns C# | unhide columns second worksheet | Cells.UnhideColumns default width | Excel to PDF conversion Aspose | .NET export Excel PDF | show hidden columns Aspose.Cells
// Common Searches: how to unhide all columns in a specific sheet with Aspose.Cells | C# Aspose.Cells unhide columns before PDF export | unhide hidden columns second worksheet Aspose.Cells | export Excel to PDF with all columns visible .NET | Cells.UnhideColumns method example
// Developer Intent: Unhide every column in the second worksheet using the default column width and then generate a PDF from the workbook.
// Use Cases: Creating printable PDFs where hidden columns must be displayed on the second sheet. | Automating batch conversion of Excel reports to PDF while ensuring full column visibility. | Preparing financial or analytical worksheets for external distribution without hidden data.
// AI Prompts: Generate C# code that unhides all columns in the second worksheet with Aspose.Cells and saves the workbook as a PDF. | Explain the parameters of Cells.UnhideColumns, especially the use of -1 for default column width. | Show how to iterate through multiple worksheets, unhide columns with default width, and combine them into a single PDF document.

using System;
using Aspose.Cells;

// Load an Excel workbook, access the second worksheet, unhide every column using the default width (‑1), and save the result as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1)
        Worksheet sheet = workbook.Worksheets[1];

        // Unhide all columns in the worksheet.
        // Using a large column count (e.g., 256) ensures all possible columns are covered.
        // Width set to -1 applies the default column width.
        sheet.Cells.UnhideColumns(0, 256, -1);

        // Save the workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
