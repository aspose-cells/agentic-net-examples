// Title: C# – Insert Column at Index 5, Preserve Formatting, Export to PDF with Aspose.Cells
// Description: Loads an Excel workbook, inserts a new column at zero‑based index 4, copies the original column's formatting to the new column using PasteOptions (PasteType.Formats), saves the changes, and converts the result to PDF via ConversionUtility.
// Keywords: Aspose.Cells insert column C# | copy column formatting Aspose.Cells | Excel to PDF conversion Aspose.Cells | PasteOptions PasteType.Formats | ConversionUtility PDF | C# Excel column manipulation
// Common Searches: how to insert a column at a specific position with Aspose.Cells | copy formatting after inserting a column in C# Excel library | convert modified Excel workbook to PDF using Aspose.Cells | Aspose.Cells PasteOptions format only copy
// Developer Intent: Add a column at position 5, duplicate the original column's style, and generate a PDF from the updated sheet.
// Use Cases: Add a placeholder column in a financial model while keeping the existing style. | Reorder columns in a reporting template, retain formatting, and produce a client‑ready PDF. | Automate invoice generation: insert a notes column, copy its formatting, and export the final document as PDF.
// AI Prompts: Generate C# code that uses Aspose.Cells to insert a column at index 5, copy the adjacent column's formatting, and save the workbook as a PDF. | Explain the role of PasteOptions with PasteType.Formats when copying column formatting in Aspose.Cells. | Provide alternative ways to convert an edited Excel workbook to PDF without using ConversionUtility.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an Excel workbook, inserts a new column at zero‑based index 4, copies the original column's formatting to the new column using PasteOptions (PasteType.Formats), saves the changes, and converts the result to PDF via ConversionUtility.
class Program
{
    static void Main()
    {
        // Paths for the original file, the intermediate modified file, and the final PDF
        string inputPath = "input.xlsx";
        string tempPath = "modified.xlsx";
        string pdfPath = "output.pdf";

        // Load the existing spreadsheet
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert a new column at position five (zero‑based index 4)
        cells.InsertColumn(4, true);

        // After insertion, the original column that was at index 4 moves to index 5.
        // Copy its formatting to the newly inserted column (index 4).
        PasteOptions pasteOptions = new PasteOptions();
        pasteOptions.PasteType = PasteType.Formats;
        cells.CopyColumns(cells, 5, 4, 1, pasteOptions);

        // Save the modified workbook to a temporary Excel file
        workbook.Save(tempPath);

        // Convert the modified Excel file to PDF using the provided conversion utility
        ConversionUtility.Convert(tempPath, pdfPath);
    }
}
