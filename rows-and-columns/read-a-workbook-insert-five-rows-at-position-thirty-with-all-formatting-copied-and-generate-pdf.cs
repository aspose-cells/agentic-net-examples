// Title: C# – Insert 5 Rows at Row 31 with Formatting and Export Workbook to PDF using Aspose.Cells
// Description: Loads an existing Excel file with Aspose.Cells, inserts five rows starting at zero‑based index 30 (row 31) while copying the style of the preceding row and updating formula references, then saves the workbook as a PDF document.
// Keywords: Aspose.Cells insert rows C# | copy row formatting Aspose.Cells | InsertRows zero based index | InsertOptions CopyFormatType SameAsAbove | UpdateReference Aspose.Cells | save Excel as PDF .NET | Aspose.Cells tutorial PDF export | add multiple rows Excel .NET | Excel to PDF conversion C#
// Common Searches: Aspose.Cells insert rows with same formatting | How to add rows at a specific position in Excel using C# | Convert modified workbook to PDF with Aspose.Cells | InsertRows with CopyFormatType.SameAsAbove example | UpdateReference effect after inserting rows Aspose.Cells
// Developer Intent: Add five blank rows at row 31, preserve the formatting of the row above, update any affected formulas, and generate a PDF version of the workbook.
// Use Cases: Expand a report template by inserting styled placeholder rows before exporting a PDF for client review. | Accommodate additional financial entries in a ledger, keep existing cell styles and formulas intact, then produce a snapshot PDF for auditors. | Create a printable schedule by inserting rows for new activities while maintaining the original layout and exporting the result as PDF.
// AI Prompts: Write C# code that uses Aspose.Cells to insert N rows at a given zero‑based index, copies the formatting from the previous row, updates references, and saves the workbook as a PDF. | Explain the role of InsertOptions.CopyFormatType.SameAsAbove and InsertOptions.UpdateReference when inserting rows with Aspose.Cells. | Provide a step‑by‑step guide for loading an Excel file, inserting rows with formatting preservation, and exporting the updated workbook to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads an existing Excel file with Aspose.Cells, inserts five rows starting at zero‑based index 30 (row 31) while copying the style of the preceding row and updating formula references, then saves the workbook as a PDF document.
class Program
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputPath = "input.xlsx";

        // Load the workbook from the file (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare insert options to copy formatting from the row above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert 5 rows starting at row index 30 (zero‑based)
        // This will push the original rows 30 and below down and copy the formatting
        worksheet.Cells.InsertRows(30, 5, insertOptions);

        // Save the modified workbook as a PDF (uses the Save(string, SaveFormat) method)
        string outputPdfPath = "output.pdf";
        workbook.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
