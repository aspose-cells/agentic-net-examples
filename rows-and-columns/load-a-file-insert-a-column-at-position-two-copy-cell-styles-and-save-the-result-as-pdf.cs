// Title: C# Example: Insert Column, Copy Formatting, and Export Excel to PDF with Aspose.Cells
// Description: Learn how to load an existing Excel file, insert a new column at the second position, copy only the formatting from the original column, and save the updated workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells insert column C# | copy column formatting Aspose.Cells | Aspose.Cells save as PDF | PasteOptions Formats Aspose.Cells | C# Excel to PDF conversion | InsertColumn method Aspose.Cells | CopyColumns formatting only | PdfSaveOptions Aspose.Cells
// Common Searches: Insert column at index 1 using Aspose.Cells C# | Copy only styles from one Excel column to another Aspose.Cells | Export modified workbook to PDF with Aspose.Cells .NET | How to use PasteOptions with PasteType.Formats in Aspose.Cells | Sample code for inserting column and saving as PDF
// Developer Intent: Add a new column at the second position, duplicate the original column’s style, and generate a PDF file.
// Use Cases: Generate PDF reports where a placeholder column inherits the layout of an existing column. | Create dynamic Excel templates that require an extra styled column before conversion to PDF. | Automate spreadsheet restructuring—adding a styled column for user input and exporting the final view as PDF.
// AI Prompts: Provide C# code that loads an Excel workbook, inserts a column at position 2, copies only the formatting from column A to the new column, and saves the result as a PDF using Aspose.Cells. | Explain the role of PasteOptions.PasteType.Formats when copying column styles in Aspose.Cells and list alternative PasteType values. | Modify the example to copy both values and formatting from the source column, then export the workbook to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Learn how to load an existing Excel file, insert a new column at the second position, copy only the formatting from the original column, and save the updated workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Input Excel file and output PDF file paths
        string inputFile = "input.xlsx";
        string outputPdf = "output.pdf";

        // Load the workbook from the existing file
        Workbook workbook = new Workbook(inputFile);

        // Get the first worksheet and its cells collection
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert a new column at position two (zero‑based index 1) and update references
        cells.InsertColumn(1, true);

        // Copy only the formatting from the original first column (index 0) to the new column (index 1)
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats
        };
        cells.CopyColumns(cells, 0, 1, 1, pasteOptions);

        // Save the modified workbook as a PDF file
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save(outputPdf, pdfOptions);
    }
}
