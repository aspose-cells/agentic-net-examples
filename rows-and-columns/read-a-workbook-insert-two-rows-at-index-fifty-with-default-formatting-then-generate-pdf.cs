// Title: C# – Insert Two Rows at Row 51 and Export Excel to PDF with Aspose.Cells
// Description: Loads an existing .xlsx file, inserts two rows at zero‑based index 50 (row 51) while keeping default formatting, and saves the updated workbook directly as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# insert rows | InsertRows method | Excel to PDF conversion | SaveFormat.Pdf | default row formatting | zero‑based row index | programmatic Excel manipulation
// Common Searches: Aspose.Cells insert rows at specific position | C# add rows to Excel and save as PDF | InsertRows method example Aspose.Cells | Convert modified workbook to PDF using Aspose.Cells | preserve formatting when inserting rows with Aspose.Cells
// Developer Intent: Add two rows at row 51 in an Excel worksheet and export the workbook as a PDF using Aspose.Cells for .NET.
// Use Cases: Insert blank rows before a summary section in a financial report, then generate a PDF for client delivery. | Adjust a template layout programmatically by adding rows and produce a print‑ready PDF. | Create a PDF version of a spreadsheet after expanding it with additional data rows.
// AI Prompts: Write C# code that uses Aspose.Cells to insert three rows at row 10 and then save the workbook as a PDF. | Show how to insert rows with custom formatting before exporting an Excel workbook to PDF using Aspose.Cells for .NET. | Provide an example that reads an existing .xlsx file, inserts rows at a given index, and converts the modified workbook to PDF.

using System;
using Aspose.Cells;

// Loads an existing .xlsx file, inserts two rows at zero‑based index 50 (row 51) while keeping default formatting, and saves the updated workbook directly as a PDF using Aspose.Cells for .NET.
class InsertRowsAndConvertToPdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFile = "input.xlsx";

        // Load the workbook (lifecycle: create via constructor with file path)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert two rows at index 50 (zero‑based). This uses the InsertRows method.
        worksheet.Cells.InsertRows(50, 2);

        // Save the modified workbook as a PDF (lifecycle: save with format)
        string outputPdf = "output.pdf";
        workbook.Save(outputPdf, SaveFormat.Pdf);
    }
}
