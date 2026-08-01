// Title: C# – Insert 3 Rows at Row 20 with Formatting and Export to PDF using Aspose.Cells
// Description: Load an existing XLSX workbook with Aspose.Cells, insert three rows starting at row 20 while preserving the original style (CopyFormatType.SameAsAbove) and updating formula references, then save the result directly as a PDF file.
// Keywords: Aspose.Cells insert rows C# | preserve formatting insert rows | CopyFormatType.SameAsAbove example | InsertRows with UpdateReference | export Excel to PDF Aspose.Cells | C# Excel row insertion PDF conversion
// Common Searches: Aspose.Cells insert rows at specific position C# | keep formatting when adding rows Aspose.Cells | save workbook as PDF after inserting rows | InsertOptions.UpdateReference effect | how to add blank rows before PDF export
// Developer Intent: Add three formatted rows at row 20 and generate a PDF from the modified workbook.
// Use Cases: Add placeholder rows for a summary section before creating a PDF report. | Expand a template with extra rows for new data while retaining existing styles. | Adjust row layout in a financial sheet and immediately produce a printable PDF.
// AI Prompts: Write C# code that uses Aspose.Cells to insert rows at a given index, copies the formatting from the row above, updates formula references, and saves the workbook as a PDF. | Explain how InsertOptions.UpdateReference influences formulas when rows are inserted with Aspose.Cells. | Show how to convert an Excel workbook to PDF after modifying its structure with Aspose.Cells.

using System;
using Aspose.Cells;

// Load an existing XLSX workbook with Aspose.Cells, insert three rows starting at row 20 while preserving the original style (CopyFormatType.SameAsAbove) and updating formula references, then save the result directly as a PDF file.
class InsertRowsAndSavePdf
{
    static void Main()
    {
        // Path to the existing spreadsheet
        string inputFile = "input.xlsx";

        // Load the workbook (create rule usage)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare insert options to preserve formatting (copy format from the row above)
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert three rows starting at row 20 (zero‑based index 19)
        worksheet.Cells.InsertRows(19, 3, insertOptions);

        // Save the modified workbook as PDF (save rule usage)
        string outputPdf = "output.pdf";
        workbook.Save(outputPdf); // format inferred from file extension

        Console.WriteLine("Rows inserted and PDF saved successfully.");
    }
}
