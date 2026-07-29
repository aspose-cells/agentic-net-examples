// Title: C# – Insert a Column, Preserve Formatting, and Export Workbook to PDF with Aspose.Cells
// Description: Loads an Excel file, adds a new column at the fifth position (zero‑based index 4), copies only the original column’s style to the inserted column using PasteOptions (PasteType.Formats), and saves the workbook as a PDF via PdfSaveOptions.
// Keywords: Aspose.Cells | C# Excel column insertion | copy column style | PasteOptions Formats | PdfSaveOptions | Excel to PDF conversion | InsertColumn method | Aspose.Cells tutorial | worksheet formatting copy | Aspose.Cells PDF export
// Common Searches: Aspose.Cells insert column at specific index C# | copy only formatting of a column Aspose.Cells | export worksheet to PDF after adding a column | use PasteOptions with PasteType.Formats in Aspose.Cells | update references when inserting a column Aspose.Cells | C# example Excel to PDF with column changes
// Developer Intent: Add a new column, duplicate its visual style, and generate a PDF file.
// Use Cases: Insert a placeholder column for user comments while keeping the original column’s look before creating a printable PDF report. | Shift existing data to make room for calculated values, retain the original formatting, and output the sheet as a PDF for distribution. | Generate a PDF version of a spreadsheet after adding a metadata column without altering any cell values.
// AI Prompts: Write C# code using Aspose.Cells that inserts a column at position 5, copies only the formatting from the adjacent column, and saves the workbook as a PDF. | Explain how PasteOptions with PasteType.Formats can be applied to copy column formatting without data in Aspose.Cells for .NET. | Provide error‑handling logic for missing input files and PDF‑save failures when inserting a column and exporting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an Excel file, adds a new column at the fifth position (zero‑based index 4), copies only the original column’s style to the inserted column using PasteOptions (PasteType.Formats), and saves the workbook as a PDF via PdfSaveOptions.
class InsertColumnAndExportPdf
{
    static void Main()
    {
        // Paths for the input Excel file and the output PDF file
        string inputPath = "input.xlsx";
        string outputPdfPath = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a new column at position five (zero‑based index 4)
        // The second parameter 'true' updates references in other worksheets
        worksheet.Cells.InsertColumn(4, true);

        // After insertion, the original column that was at index 4 is now at index 5.
        // Copy its formatting to the newly inserted column (index 4).
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats   // Copy only formatting, not data
        };
        worksheet.Cells.CopyColumns(
            worksheet.Cells,   // source cells collection
            5,                 // source column index (original column shifted right)
            4,                 // destination column index (newly inserted column)
            1,                 // number of columns to copy
            pasteOptions);    // copy only formats

        // Export the modified workbook to PDF
        workbook.Save(outputPdfPath, new PdfSaveOptions());

        Console.WriteLine("Column inserted, formatting copied, and PDF saved successfully.");
    }
}
