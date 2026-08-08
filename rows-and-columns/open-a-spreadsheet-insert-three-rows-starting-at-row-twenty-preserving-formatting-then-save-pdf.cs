// Title: C# – Insert Rows at Row 20 with Formatting and Save as PDF using Aspose.Cells
// Description: Loads an existing XLSX file, inserts three rows beginning at row 20 while copying the style of the row above, updates references, and saves the workbook as a PDF document.
// Keywords: Aspose.Cells C# insert rows | preserve row formatting Aspose.Cells | CopyFormatType SameAsAbove | export Excel to PDF C# | InsertRows method Aspose.Cells | global spreadsheet automation | GitHub Aspose.Cells example
// Common Searches: How to insert rows at a specific index with Aspose.Cells for .NET | Insert rows and keep formatting in C# using Aspose.Cells | Save an Excel workbook as PDF after adding rows | Aspose.Cells InsertRows row 20 example | Copy formatting from above row when inserting rows
// Developer Intent: Add three rows at row 20, retain the original styling, and export the worksheet to PDF.
// Use Cases: Add blank rows for a new data section in a report template before generating a PDF. | Programmatically expand a spreadsheet for printing while preserving existing styles. | Maintain cell formats during row insertion and produce a PDF for stakeholder distribution.
// AI Prompts: Write C# code with Aspose.Cells that inserts N rows at a given zero‑based index, copies the format from the row above, and saves the workbook as a PDF. | Explain the effect of InsertOptions.CopyFormatType.SameAsAbove on row insertion and the resulting PDF output. | Create a variant that inserts rows without updating formulas and then converts the workbook to PDF.

using System;
using Aspose.Cells;

// Loads an existing XLSX file, inserts three rows beginning at row 20 while copying the style of the row above, updates references, and saves the workbook as a PDF document.
class InsertRowsAndSavePdf
{
    static void Main()
    {
        // Path to the existing spreadsheet
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputPdf = "output.pdf";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare insert options to preserve formatting (copy format from the row above)
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert three rows starting at row 20 (zero‑based index 19)
        worksheet.Cells.InsertRows(19, 3, insertOptions);

        // Save the modified workbook as a PDF document
        workbook.Save(outputPdf, SaveFormat.Pdf);
    }
}
