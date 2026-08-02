// Title: Insert Column at Index 5, Copy Only Formatting, and Export to PDF with Aspose.Cells for .NET (C#)
// Description: Shows how to load an Excel workbook using Aspose.Cells, insert a new column at zero‑based index 4 (the fifth column), copy only the formatting from the original column to the new one with PasteOptions, save the workbook to a temporary file, and convert it to PDF via ConversionUtility.
// Keywords: Aspose.Cells insert column C# | copy column formatting Aspose.Cells | Aspose.Cells PDF conversion | C# Excel column manipulation | PasteOptions Formats | ConversionUtility Aspose.Cells | temporary workbook save | zero based column index | .NET Aspose.Cells
// Common Searches: Aspose.Cells insert column at specific position C# | Copy only formats of a column with Aspose.Cells | Convert modified Excel workbook to PDF using Aspose.Cells | How to use PasteOptions PasteType.Formats in Aspose.Cells | Insert column and preserve style before PDF export
// Developer Intent: Add a new column at the fifth position, duplicate only its formatting, and generate a PDF from the updated workbook.
// Use Cases: Create a placeholder column in a financial report template while keeping the original style before publishing as PDF. | Programmatically adjust column layout in a large dataset and deliver a printable PDF to clients. | Automate invoice PDF generation by inserting a notes column without altering existing formatting. | Prepare regulatory filing spreadsheets by adding columns for new data fields, preserving formatting, then exporting to PDF.
// AI Prompts: Generate C# code using Aspose.Cells to insert a column at index 4, copy only the formatting from column 5 to the new column, and save the workbook as a PDF. | Explain the role of PasteOptions.PasteType.Formats when copying column styles in Aspose.Cells and demonstrate PDF conversion with ConversionUtility. | Provide a step‑by‑step guide for inserting a column, copying formats, and exporting to PDF in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Shows how to load an Excel workbook using Aspose.Cells, insert a new column at zero‑based index 4 (the fifth column), copy only the formatting from the original column to the new one with PasteOptions, save the workbook to a temporary file, and convert it to PDF via ConversionUtility.
class InsertColumnAndExportPdf
{
    static void Main()
    {
        // Paths for the original workbook, a temporary modified workbook, and the final PDF
        string inputPath = "input.xlsx";
        string tempPath = "temp_modified.xlsx";
        string outputPdf = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Work with the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert a new column at position five (zero‑based index 4)
        cells.InsertColumn(4);

        // Copy only the formatting from the original column (now at index 5) to the new column (index 4)
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats   // copy formats only
        };
        cells.CopyColumns(cells, 5, 4, 1, pasteOptions);

        // Save the modified workbook to a temporary file
        workbook.Save(tempPath);

        // Convert the temporary workbook to PDF using the provided conversion utility
        ConversionUtility.Convert(tempPath, outputPdf);

        // Optional: clean up the temporary file
        // System.IO.File.Delete(tempPath);
    }
}
