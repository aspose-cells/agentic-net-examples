// Title: Insert Column, Duplicate Formatting, and Export to PDF with Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel workbook, inserts a new column at index 1, copies the original first column’s values and styles to the new column, and saves the result as a PDF using Aspose.Cells.
// Keywords: Aspose.Cells | C# | insert column | copy column formatting | Excel to PDF | PdfSaveOptions | Workbook.Save | InsertColumn method | CopyColumn method
// Common Searches: Aspose.Cells insert column at specific index | Copy column formatting Aspose.Cells C# | Export Excel workbook to PDF with Aspose.Cells | Preserve formulas when inserting columns Aspose.Cells | Shift references after column insertion Aspose.Cells
// Developer Intent: Add a column at position two, replicate the first column’s data and style, then generate a PDF of the workbook.
// Use Cases: Create a standardized report layout by inserting a blank column next to existing data and copying its formatting. | Automate spreadsheet restructuring while maintaining formula references before distributing a PDF version. | Generate printable PDFs from Excel files after programmatically adjusting column order or adding placeholders.
// AI Prompts: Write C# code using Aspose.Cells to insert a column at index 1, copy the first column’s content and formatting, and save the workbook as a PDF. | Explain how the second parameter of InsertColumn controls reference updating and how to keep formulas intact when copying columns. | Provide robust error handling for missing input files, permission issues, and PDF conversion failures in an Aspose.Cells workflow.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// C# example that loads an Excel workbook, inserts a new column at index 1, copies the original first column’s values and styles to the new column, and saves the result as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert a new column at position two (index 1) and update references
        cells.InsertColumn(1, true);

        // Copy data and formatting from the original first column (index 0) to the new column (index 1)
        cells.CopyColumn(cells, 0, 1);

        // Save the modified workbook as a PDF file
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("output.pdf", pdfOptions);
    }
}
