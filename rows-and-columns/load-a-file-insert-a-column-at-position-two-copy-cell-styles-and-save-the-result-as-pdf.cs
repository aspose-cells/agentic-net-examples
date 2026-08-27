// Title: Insert a column at position two, copy its formatting from the first column, and save the workbook as PDF using Aspose.Cells for .NET
// AI Prompts: Insert a column at index 1, copy just the formats from column A, and generate a PDF file with Aspose.Cells in C#. | Using Aspose.Cells for .NET, add a second column, apply PasteOptions to transfer formatting, then save the workbook as PDF.
// Common Searches: C# Aspose.Cells insert column and preserve formatting | How to copy only cell styles when adding a new column with Aspose.Cells | Export Excel to PDF after modifying columns using Aspose.Cells .NET | Maintain formula references after inserting a column in Aspose.Cells
// Tags: insert column with PasteOptions formats Aspose.Cells | copy column formatting Aspose.Cells C# | save workbook as PDF Aspose.Cells | update formulas after column insertion Aspose.Cells | PdfSaveOptions configuration Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// The example loads 'input.xlsx', inserts a new column at position two while updating references, copies only the formatting from the original first column to the new column using PasteOptions with PasteType.Formats, and saves the result as 'output.pdf' via PdfSaveOptions.
class InsertColumnAndSavePdf
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert a new column at position two (index 1) and update references
        cells.InsertColumn(1, true);

        // Copy only the formatting (styles) from the original column (index 0) to the new column (index 1)
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats
        };
        cells.CopyColumns(cells, 0, 1, 1, pasteOptions);

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", new PdfSaveOptions());
    }
}
