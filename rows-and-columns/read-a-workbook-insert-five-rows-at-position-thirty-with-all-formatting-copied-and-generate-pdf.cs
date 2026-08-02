// Title: Insert 5 Rows with Copied Formatting at Row 31 and Save as PDF – Aspose.Cells C# Example
// Description: Loads an existing Excel workbook, inserts five rows beginning at row 31 while copying the formatting of the rows above and updating formula references, then saves the modified file directly as a PDF.
// Keywords: Aspose.Cells InsertRows | CopyFormatType.SameAsAbove | C# Excel row insertion | preserve cell formatting Aspose.Cells | update references after insert | save workbook as PDF | Aspose.Cells PDF conversion | InsertOptions C# | Excel row 31 insertion .NET | Aspose.Cells export to PDF
// Common Searches: Aspose.Cells insert rows with same formatting | How to add rows at a specific position in Excel using C# | Copy formatting from above rows Aspose.Cells | Export modified workbook to PDF with Aspose.Cells | Update formulas after inserting rows Aspose.Cells
// Developer Intent: Add five rows at row 31 while preserving formatting, then generate a PDF of the workbook.
// Use Cases: Add blank rows for a new data section in a template, keep header styles unchanged, and produce a printable PDF report. | Insert rows in a financial sheet, maintain existing formula links, and export the updated worksheet as a PDF for client distribution. | Create an invoice PDF by inserting extra line‑item rows, preserving the original cell styles and formulas.
// AI Prompts: Show C# code that uses Aspose.Cells InsertOptions to insert five rows at row 31 with the same formatting as the rows above and then saves the workbook as a PDF. | Explain how CopyFormatType.SameAsAbove works when inserting rows and how it affects formula references in Aspose.Cells. | Provide a step‑by‑step guide to insert rows at a specific index, copy formatting, update references, and convert the workbook to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads an existing Excel workbook, inserts five rows beginning at row 31 while copying the formatting of the rows above and updating formula references, then saves the modified file directly as a PDF.
class Program
{
    static void Main()
    {
        // Load the existing workbook from disk
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure insertion options to copy formatting from the rows above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert five rows starting at row index 30 (Excel row 31, zero‑based indexing)
        worksheet.Cells.InsertRows(30, 5, insertOptions);

        // Save the modified workbook as a PDF file
        string outputFile = "output.pdf";
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}
